"""Remplit le document de suivi de tournoi à partir des données start.gg,
avec la mise en page de vierge.py.

Usage :
    python cast-document.py            (le script pose ses questions)
    python cast-document.py <URL de l'event start.gg> [--sortie fichier.docx]

Le token API start.gg est lu dans le fichier .env (STARTGG_TOKEN=...).

Les personnages sont déduits des games saisies sur start.gg : on parcourt
l'historique de sets de chaque joueur (plusieurs pages si besoin) ; si aucun
TO n'a saisi les persos, la case reste vide.

Fichiers optionnels (2 colonnes, séparateur « ; » ou « , ») :
    rankings.csv : pseudo ; classement   (ex. KID;LUMI #87)
    teams.csv    : tricode ; nom complet (ex. TLSE;Toulouse Last Stock Esport)

Les réponses de l'API sont gardées dans cache-startgg.json : supprimer ce
fichier pour forcer un nouveau téléchargement.
"""

import argparse
import csv
import json
import os
import re
import sys
import time
import unicodedata
import urllib.error
import urllib.request
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.image.exceptions import UnrecognizedImageError
from docx.shared import Cm, Pt, RGBColor

import vierge
from vierge import (
    GRIS, emplacement_image, fond_cellule, lien_hypertexte, nota_bene, saut_de_page,
    source, surlignage, tableau, titre,
)

# --- Configuration ----------------------------------------------------------

API_URL = "https://api.start.gg/gql/alpha"
ID_ULTIMATE = 1386
PAUSE_ENTRE_REQUETES = 0.8  # limite start.gg : 80 requêtes / minute
FICHIER_CACHE = "cache-startgg.json"

SETS_PAR_PAGE = 15
PAGES_MAX_PERSOS = 5   # pages de sets parcourues au maximum pour trouver les persos
PARTIES_MIN = 10       # on arrête de chercher dès qu'on a autant de games renseignées
PAGES_MAX_H2H = 10     # historique plus profond pour les joueurs du Top 8
PART_MIN_PERSO = 0.3   # un perso est retenu s'il fait au moins 30 % des games du main
NB_PERSOS_MAX = 3
PERSOS_IGNORES = {"Random Character", "Random"}

DOSSIER_ICONES = Path("Pngs-20261001T134705Z-1-001/Pngs")
VERT = "00B050"
GRIS_CLAIR = "D9D9D9"
COULEURS_TIERS = ["FF7F7F", "FFBF7F", "FFDF7F", "FFFF7F", "BFFF7F", "7FFF7F", "7FFFFF"]
TYPES_BRACKET = {"DOUBLE_ELIMINATION": "double élimination", "SINGLE_ELIMINATION": "simple élimination",
                 "ROUND_ROBIN": "round robin", "SWISS": "rondes suisses"}

# Noms start.gg dont l'icône porte un autre nom de fichier
ICONES_SPECIALES = {
    "hero": "dq_hero", "incineroar": "gaogaen", "pyra": "pyra_and_mythra",
    "mythra": "pyra_and_mythra", "min_min": "minmin", "piranha_plant": "packun_flower",
    "banjo_kazooie": "banjo_and_kazooie", "rosalina": "rosalina_and_luma", "r_o_b": "rob",
    "game_and_watch": "mr_game_and_watch", "mii_swordfighter": "mii_sword_fighter",
    "mii_swordsman": "mii_sword_fighter", "alex": "steve",
}

# --- Requêtes GraphQL -------------------------------------------------------

REQUETE_EVENT = """
query Evenement($slug: String) {
  event(slug: $slug) {
    id name slug numEntrants startAt
    tournament { name venueName venueAddress streams { streamName streamSource } }
    phases { id name phaseOrder bracketType numSeeds }
    phaseGroups { id displayIdentifier phase { id } }
  }
}
"""

REQUETE_SEEDS = """
query Seeds($id: ID!, $page: Int!) {
  phaseGroup(id: $id) {
    seeds(query: {page: $page, perPage: 64}) {
      pageInfo { totalPages }
      nodes {
        seedNum
        entrant {
          participants {
            prefix gamerTag
            player { id }
            user { slug location { city state country } }
          }
        }
      }
    }
  }
}
"""

REQUETE_INSCRITS = """
query Inscrits($id: ID!, $page: Int!) {
  event(id: $id) {
    entrants(query: {page: $page, perPage: 64}) {
      pageInfo { totalPages }
      nodes {
        initialSeedNum
        participants {
          prefix gamerTag
          player { id }
          user { slug location { city state country } }
        }
      }
    }
  }
}
"""

REQUETE_SETS = """
query Sets($id: ID!, $page: Int!, $perPage: Int!) {
  player(id: $id) {
    sets(page: $page, perPage: $perPage) {
      nodes {
        fullRoundText
        event { slug videogame { id } tournament { name } }
        slots {
          entrant { id participants { player { id } } }
          standing { stats { score { value } } }
        }
        games { selections { entrant { id } character { name } } }
      }
    }
  }
}
"""


# --- Accès à l'API start.gg -------------------------------------------------

class StartGG:
    def __init__(self, token, chemin_cache=FICHIER_CACHE):
        self.token = token
        self.chemin_cache = Path(chemin_cache)
        self.cache = json.loads(self.chemin_cache.read_text("utf-8")) if self.chemin_cache.exists() else {}

    def requete(self, query, variables):
        corps = json.dumps({"query": query, "variables": variables}).encode()
        for essai in range(5):
            time.sleep(PAUSE_ENTRE_REQUETES)
            req = urllib.request.Request(API_URL, data=corps, headers={
                "Authorization": f"Bearer {self.token}", "Content-Type": "application/json",
            })
            try:
                with urllib.request.urlopen(req, timeout=30) as reponse:
                    donnees = json.load(reponse)
            except urllib.error.HTTPError as erreur:
                if erreur.code == 429 or erreur.code >= 500:
                    time.sleep(15 * (essai + 1))
                    continue
                raise
            if donnees.get("data"):
                return donnees["data"]
            message = str(donnees.get("errors") or donnees.get("message"))
            if "rate limit" in message.lower():
                time.sleep(15 * (essai + 1))
                continue
            raise RuntimeError(f"Erreur start.gg : {message}")
        raise RuntimeError("start.gg ne répond pas (limite de requêtes ?)")

    def sauver_cache(self):
        self.chemin_cache.write_text(json.dumps(self.cache, ensure_ascii=False), "utf-8")

    def evenement(self, slug):
        return self.requete(REQUETE_EVENT, {"slug": slug})["event"]

    def seeds(self, phase_group_id):
        noeuds, page = [], 1
        while True:
            seeds = self.requete(REQUETE_SEEDS, {"id": phase_group_id, "page": page})["phaseGroup"]["seeds"]
            noeuds += seeds["nodes"]
            if page >= seeds["pageInfo"]["totalPages"]:
                return noeuds
            page += 1

    def inscrits(self, event_id):
        noeuds, page = [], 1
        while True:
            entrants = self.requete(REQUETE_INSCRITS, {"id": event_id, "page": page})["event"]["entrants"]
            noeuds += entrants["nodes"]
            if page >= entrants["pageInfo"]["totalPages"]:
                return noeuds
            page += 1

    def sets_joueur(self, player_id, nb_pages):
        """Renvoie les sets du joueur (les plus récents d'abord) sur au moins
        nb_pages pages, en complétant le cache si nécessaire."""
        entree = self.cache.setdefault(f"sets:{player_id}", {"pages": 0, "fin": False, "sets": []})
        while entree["pages"] < nb_pages and not entree["fin"]:
            donnees = self.requete(REQUETE_SETS, {
                "id": player_id, "page": entree["pages"] + 1, "perPage": SETS_PAR_PAGE,
            })
            noeuds = ((donnees.get("player") or {}).get("sets") or {}).get("nodes") or []
            entree["sets"] += noeuds
            entree["pages"] += 1
            entree["fin"] = len(noeuds) < SETS_PAR_PAGE
            self.sauver_cache()
        return entree["sets"]


# --- Analyse des données ----------------------------------------------------

def charger_env(chemin=".env"):
    """Charge les lignes CLE=valeur du fichier .env dans os.environ."""
    chemin = Path(chemin)
    if not chemin.exists():
        return
    for ligne in chemin.read_text("utf-8").splitlines():
        cle, egal, valeur = ligne.partition("=")
        if egal and not cle.strip().startswith("#"):
            os.environ.setdefault(cle.strip(), valeur.strip().strip('"').strip("'"))


def normaliser_slug(texte):
    """Extrait « tournament/<tournoi>/event/<event> » d'une URL start.gg (None si absent)."""
    trouve = re.search(r"tournament/[^/?#\s]+/event/[^/?#\s]+", texte)
    return trouve.group(0) if trouve else None


def demander(question, exemple="", defaut=""):
    """Pose une question dans le terminal ; Entrée seule renvoie la valeur par défaut."""
    if exemple:
        print(f"  (ex. {exemple})")
    suffixe = f" [{defaut}]" if defaut else ""
    return input(f"{question}{suffixe} : ").strip() or defaut


def confirmer(question):
    return demander(f"{question} (O/n)").lower() in ("", "o", "oui", "y", "yes")


def demander_evenement(api, slug_initial=None):
    """Demande l'URL de l'event jusqu'à ce qu'elle soit valide et confirmée."""
    slug = slug_initial
    while True:
        if not slug:
            reponse = demander(
                "\nURL de l'event start.gg",
                "https://www.start.gg/tournament/pivot-grab-1-main-tendue-31/event/pivot-grab-1v1",
            )
            slug = normaliser_slug(reponse)
            if not slug:
                print("  URL invalide : elle doit contenir « tournament/<tournoi>/event/<event> ».")
                print("  Ouvre l'event sur start.gg (onglet de l'event, pas seulement le tournoi) et copie l'URL.")
                continue
        print("Recherche de l'event sur start.gg...")
        evenement = api.evenement(slug)
        if not evenement:
            print(f"  Aucun event trouvé pour « {slug} ».")
        else:
            date = datetime.fromtimestamp(evenement["startAt"]).strftime("%d/%m/%Y") if evenement.get("startAt") else "?"
            print(f"  Trouvé : {evenement['tournament']['name']} – {evenement['name']} "
                  f"({evenement['numEntrants']} entrants, le {date})")
            if confirmer("  Est-ce le bon event ?"):
                return evenement
        slug = None


def demander_infos():
    """Questions sur les champs que start.gg ne fournit pas (vide = à compléter à la main)."""
    print("\nInformations complémentaires (Entrée pour laisser le champ à compléter plus tard)")
    return {
        "organisation": demander("Organisation de l'évènement", "Association G.R.A.B. & TOULOUSE LAST STOCK"),
        "participation": demander("Avec la participation de", "Main Tendue 31"),
        "format_sets": demander("Format des sets", "FULL BO5, ou BO3 puis BO5 à partir du Top 8"),
        "supermajor": demander("Lien Supermajor de l'event", "https://www.supermajor.gg/ultimate/event/S1423245"),
        "auteur": demander("Création & conception du document", "Loïc « loicsamaaa » RAKOTOSON"),
    }


def lire_csv(chemin):
    """Lit un CSV à 2 colonnes en dictionnaire {colonne 1 en minuscules: colonne 2}."""
    chemin = Path(chemin)
    if not chemin.exists():
        return {}
    texte = chemin.read_text("utf-8-sig")
    separateur = ";" if ";" in texte.splitlines()[0] else ","
    return {
        ligne[0].strip().lower(): ligne[1].strip()
        for ligne in csv.reader(texte.splitlines(), delimiter=separateur)
        if len(ligne) >= 2 and ligne[0].strip()
    }


def premiere_phase(evenement):
    """Phase d'entrée (les pools) : celle qui contient le plus de joueurs.
    phaseOrder n'est pas fiable, certains TO numérotent les phases à l'envers."""
    phases = evenement.get("phases") or []
    if not phases:
        return None
    return max(phases, key=lambda phase: (phase.get("numSeeds") or 0, -phase["phaseOrder"]))


def creer_joueur(entrant, seed):
    """Construit un joueur à partir d'un entrant start.gg (None si l'entrant est vide)."""
    participant = ((entrant or {}).get("participants") or [None])[0]
    if not participant:
        return None
    utilisateur = participant.get("user") or {}
    lieu = utilisateur.get("location") or {}
    ville = (lieu.get("city") or lieu.get("state") or "").strip()
    return {
        "pseudo": participant["gamerTag"],
        "tag": re.sub(r"[\s|]+$", "", participant.get("prefix") or ""),
        "seed": seed,
        "ville": ville.title() if ville.isupper() else ville,
        "player_id": str(participant["player"]["id"]),
        "url": f"https://www.start.gg/{utilisateur['slug']}" if utilisateur.get("slug") else "",
    }


def joueurs_par_pool(api, evenement):
    """Renvoie [(titre, [joueurs])] : un tableau par pool de la première phase,
    ou la liste des inscrits si les pools ne sont pas encore publiés."""
    phase = premiere_phase(evenement)
    groupes = [g for g in evenement.get("phaseGroups") or [] if phase and g["phase"]["id"] == phase["id"]]
    groupes.sort(key=lambda g: (len(g["displayIdentifier"]), g["displayIdentifier"]))

    pools = []
    for groupe in groupes:
        joueurs = [creer_joueur(seed.get("entrant"), seed["seedNum"]) for seed in api.seeds(groupe["id"])]
        pools.append((f"Pool {groupe['displayIdentifier']}", [j for j in joueurs if j]))
    if any(joueurs for _, joueurs in pools):
        return [(nom, joueurs) for nom, joueurs in pools if joueurs]

    print("  Les pools ne sont pas encore remplis sur start.gg : le document utilisera la liste des inscrits.")
    joueurs = [creer_joueur(entrant, entrant.get("initialSeedNum")) for entrant in api.inscrits(evenement["id"])]
    return [("Liste des inscrits", [j for j in joueurs if j])]


def analyser_sets(sets, player_id, slug_exclu=None):
    """Renvoie (Counter des persos joués, liste des sets 1v1 Ultimate)."""
    persos = Counter()
    resultats = []
    for s in sets:
        event = s.get("event") or {}
        if int((event.get("videogame") or {}).get("id") or 0) != ID_ULTIMATE:
            continue
        slots = [slot for slot in s.get("slots") or [] if slot.get("entrant")]
        if len(slots) != 2:
            continue
        participants = [slot["entrant"].get("participants") or [] for slot in slots]
        if any(len(p) != 1 for p in participants):  # 2v2 : les persos du partenaire fausseraient le compte
            continue
        ids = [str((p[0].get("player") or {}).get("id")) for p in participants]
        if player_id not in ids:
            continue
        moi, adversaire = (slots[0], slots[1]) if ids[0] == player_id else (slots[1], slots[0])

        for game in s.get("games") or []:
            for selection in game.get("selections") or []:
                if (selection.get("character") and selection.get("entrant")
                        and selection["character"]["name"] not in PERSOS_IGNORES
                        and str(selection["entrant"]["id"]) == str(moi["entrant"]["id"])):
                    persos[selection["character"]["name"]] += 1

        def score(slot):
            return ((((slot.get("standing") or {}).get("stats") or {}).get("score")) or {}).get("value")

        score_moi, score_adversaire = score(moi), score(adversaire)
        # Les DQ sont notées -1 : on les ignore pour les head-to-heads
        if (score_moi is None or score_adversaire is None or min(score_moi, score_adversaire) < 0
                or event.get("slug") == slug_exclu):
            continue
        resultats.append({
            "adversaire": ids[1] if ids[0] == player_id else ids[0],
            "score": int(score_moi),
            "score_adversaire": int(score_adversaire),
            "tournoi": (event.get("tournament") or {}).get("name", ""),
            "round": s.get("fullRoundText") or "",
        })
    return persos, resultats


def persos_principaux(api, joueur):
    """Parcourt l'historique du joueur jusqu'à trouver assez de games renseignées."""
    persos = Counter()
    for nb_pages in range(1, PAGES_MAX_PERSOS + 1):
        persos, _ = analyser_sets(api.sets_joueur(joueur["player_id"], nb_pages), joueur["player_id"])
        if sum(persos.values()) >= PARTIES_MIN:
            break
    if not persos:
        return []
    maximum = persos.most_common(1)[0][1]
    return [nom for nom, n in persos.most_common(NB_PERSOS_MAX) if n >= maximum * PART_MIN_PERSO]


def cle_icone(nom):
    """« Pyra & Mythra », « pyra & mythra » et « pyra_and_mythra » donnent la même clé."""
    cle = unicodedata.normalize("NFKD", nom).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "_", cle.replace("&", "and")).strip("_")


def chemin_icone(nom_perso):
    if not hasattr(chemin_icone, "index"):
        chemin_icone.index = {cle_icone(p.stem): p for p in DOSSIER_ICONES.glob("*.png")}
    cle = cle_icone(nom_perso)
    return chemin_icone.index.get(cle) or chemin_icone.index.get(ICONES_SPECIALES.get(cle, cle))


# --- Helpers de mise en page ------------------------------------------------

def ecrire(cellule, texte, gras=False, italique=False, alignement=None):
    paragraphe = cellule.paragraphs[0]
    if alignement is not None:
        paragraphe.alignment = alignement
    run = paragraphe.add_run(texte or "")
    run.bold = gras
    run.italic = italique
    return run


def icones(paragraphe, persos, hauteur_cm=0.5):
    for nom in persos:
        chemin = chemin_icone(nom)
        if chemin:
            try:
                paragraphe.add_run().add_picture(str(chemin), height=Cm(hauteur_cm))
                continue
            except UnrecognizedImageError:  # ex. fichier WebP renommé en .png
                pass
        paragraphe.add_run(f"{nom} ").font.size = Pt(8)


def ligne_lien(doc, libelle, url, gras_italique=False):
    paragraphe = doc.add_paragraph()
    run = paragraphe.add_run(libelle)
    run.bold = run.italic = gras_italique
    paragraphe.add_run(" : ")
    if url:
        lien_hypertexte(paragraphe, url, url.removeprefix("https://"))
    else:
        paragraphe.add_run("……………………………………………………")


def tableau_tiers(doc, compte, avec_icones, regrouper_uniques=False):
    """Tableau façon tier list : une ligne par nombre d'occurrences. Avec
    regrouper_uniques, les entrées présentes une seule fois forment un dernier
    palier « Autre »."""
    groupes = defaultdict(list)
    for nom, n in compte.items():
        groupes[n].append(nom)
    autres = groupes.pop(1, []) if regrouper_uniques else []
    if not groupes and not autres:
        doc.add_paragraph("Aucune donnée disponible sur start.gg.").runs[0].italic = True
        return
    lignes = [(str(n), noms) for n, noms in sorted(groupes.items(), reverse=True)]
    if autres:
        lignes.append(("Autre", autres))
    table = doc.add_table(rows=len(lignes), cols=2)
    table.style = "Table Grid"
    for i, (libelle, noms) in enumerate(lignes):
        gauche, droite = table.rows[i].cells
        gauche.width, droite.width = Cm(1.5), Cm(14.5)
        couleur = GRIS_CLAIR if libelle == "Autre" else COULEURS_TIERS[min(i, len(COULEURS_TIERS) - 1)]
        fond_cellule(gauche, couleur)
        ecrire(gauche, libelle, gras=True, alignement=WD_ALIGN_PARAGRAPH.CENTER)
        if avec_icones:
            icones(droite.paragraphs[0], sorted(noms), hauteur_cm=0.8)
        elif libelle == "Autre":
            nb_joueurs = sum(compte[nom] for nom in noms)
            ecrire(droite, f"{nb_joueurs} joueurs", gras=True)
            detail = droite.add_paragraph(f"Villes représentées par un seul joueur : {', '.join(sorted(noms, key=str.lower))}")
            detail.paragraph_format.space_after = Pt(0)
            detail.runs[0].font.size = Pt(8)
            detail.runs[0].italic = True
        else:
            ecrire(droite, ", ".join(sorted(noms, key=str.lower)))


def harmoniser_villes(joueurs):
    """Regroupe les orthographes d'une même ville (« lyon », « LYON », « Lyon » → « Lyon »)."""
    variantes = defaultdict(Counter)
    for joueur in joueurs:
        if joueur["ville"]:
            variantes[joueur["ville"].casefold()][joueur["ville"]] += 1

    def forme_retenue(compte):
        bien_ecrites = [v for v in compte if not v.islower() and not v.isupper()]
        if bien_ecrites:
            return max(bien_ecrites, key=compte.get)
        ville = compte.most_common(1)[0][0]
        return ville.title() if ville.isupper() else ville[0].upper() + ville[1:]

    noms = {cle: forme_retenue(compte) for cle, compte in variantes.items()}
    for joueur in joueurs:
        if joueur["ville"]:
            joueur["ville"] = noms[joueur["ville"].casefold()]


# --- Sections du document ---------------------------------------------------

def informations_generales(doc, evenement, infos):
    tournoi = evenement["tournament"]
    titre(doc, "Informations générales")
    vierge.ligne_champ(doc, "Organisation", infos["organisation"], gras_italique=True)
    doc.add_paragraph(f"avec la participation de {infos['participation'] or '……………………'}", style="List Bullet")
    lieu = " - ".join(x for x in (tournoi.get("venueName"), tournoi.get("venueAddress")) if x)
    vierge.ligne_champ(doc, "Lieu de l’évènement", lieu)
    if evenement.get("startAt"):
        vierge.ligne_champ(doc, "Date", datetime.fromtimestamp(evenement["startAt"]).strftime("%d/%m/%Y"))
    vierge.ligne_champ(doc, f"Nombre de joueurs/euses ({evenement['name']})", f"{evenement['numEntrants']} entrants")

    paragraphe = doc.add_paragraph("Diffusions : ")
    streams = [s for s in tournoi.get("streams") or [] if s.get("streamName")]
    for i, stream in enumerate(streams):
        if i:
            paragraphe.add_run(", ")
        if stream.get("streamSource") == "TWITCH":
            url = f"https://twitch.tv/{stream['streamName']}"
            lien_hypertexte(paragraphe, url, url.removeprefix("https://"))
        else:
            paragraphe.add_run(stream["streamName"])
    if not streams:
        paragraphe.add_run("……………………………………………………")

    ligne_lien(doc, "Bracket - start.gg", f"https://www.start.gg/{evenement['slug']}", gras_italique=True)
    ligne_lien(doc, "Profil Supermajor", infos["supermajor"], gras_italique=True)
    ligne_lien(doc, "UltRank Tournament Tierer - @kenniky1", "https://kenniky.github.io/web_ultrank_scorer.html",
               gras_italique=True)
    doc.add_paragraph()
    emplacement_image(doc, "Capture UltRank Tournament Tierer", 7)
    source(doc, f"UltRank Tournament Tierer – {vierge.TOURNOI}")
    saut_de_page(doc)


def formats_et_ruleset(doc, evenement, format_sets):
    phase = premiere_phase(evenement) or {}
    titre(doc, "Formats des brackets et ruleset")
    paragraphe = doc.add_paragraph()
    paragraphe.add_run("Format du tournoi et des sets").bold = True
    type_bracket = TYPES_BRACKET.get(phase.get("bracketType"), "……………… élimination")
    paragraphe.add_run(f" : Bracket à {type_bracket} - ")
    run = paragraphe.add_run(format_sets or "BO…")
    run.bold = run.underline = True
    run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    emplacement_image(doc, "Capture des brackets / pools (start.gg)", 7)
    emplacement_image(doc, "Image du ruleset (stages, règles de ban)", 10)
    source(doc, f"start.gg – {vierge.TOURNOI}")
    vierge.ligne_champ(doc, "Liens utiles")
    saut_de_page(doc)


def profils_joueurs(doc, pools, rankings, teams):
    titre(doc, "Profil des joueurs/euses")
    entetes = ["Tag/Struc.*", "Pseudo - start.gg", "Seed", "Région/Ville", "Perso(s)", "PR/Rank*"]
    largeurs = [2.6, 4.4, 1.3, 3.0, 2.2, 2.8]
    for titre_pool, joueurs in pools:
        titre(doc, titre_pool, niveau=2, italique=True)
        table = tableau(doc, entetes, len(joueurs), largeurs)
        # Seed croissant ; les joueurs sans seed à la fin, par ordre alphabétique
        par_seed = sorted(joueurs, key=lambda j: (j["seed"] is None, j["seed"] or 0, j["pseudo"].lower()))
        for ligne, joueur in zip(table.rows[1:], par_seed):
            cellules = ligne.cells
            ecrire(cellules[0], joueur["tag"])
            if joueur["url"]:
                lien_hypertexte(cellules[1].paragraphs[0], joueur["url"], joueur["pseudo"])
            else:
                ecrire(cellules[1], joueur["pseudo"])
            ecrire(cellules[2], str(joueur["seed"] or ""), gras=True, alignement=WD_ALIGN_PARAGRAPH.RIGHT)
            ecrire(cellules[3], joueur["ville"])
            cellules[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            icones(cellules[4].paragraphs[0], joueur["persos"])
            rang = rankings.get(joueur["pseudo"].lower())
            if rang:
                surlignage(ecrire(cellules[5], rang, gras=True))
            else:
                fond_cellule(cellules[5], GRIS)
        saut_de_page(doc)

    titre(doc, "Abréviations et annotations", niveau=3)
    tous = [j for _, joueurs in pools for j in joueurs]
    tricodes = sorted({t for j in tous for t in re.split(r"[|\s]+", j["tag"]) if re.search(r"\w", t)},
                      key=str.lower)
    moitie = (len(tricodes) + 1) // 2
    table = tableau(doc, ["Abré.°/Tricode", "Nom complet", "Abré.°/Tricode", "Nom complet"], moitie,
                    [3.0, 5.2, 3.0, 5.2])
    for i, tricode in enumerate(tricodes):
        cellules = table.rows[1 + i % moitie].cells
        colonne = 0 if i < moitie else 2
        ecrire(cellules[colonne], tricode)
        nom = teams.get(tricode.lower())
        if nom:
            ecrire(cellules[colonne + 1], nom, gras=True, italique=True)
        else:
            fond_cellule(cellules[colonne + 1], GRIS)

    doc.add_paragraph()
    nota_bene(
        doc,
        "les classements inscrits dans les tableaux (colonne « PR/Rank »), ci-dessus, "
        "sont issus des derniers Power Rankings en date.",
        surligne=True,
    ).add_run(" (……………………)")
    table = tableau(doc, ["Abréviations", "Nom complet", "Abréviations", "Nom complet"], len(vierge.RANKINGS),
                    [2.6, 5.6, 2.6, 5.6])
    for ligne, valeurs in zip(table.rows[1:], vierge.RANKINGS):
        for i, valeur in enumerate(valeurs):
            ecrire(ligne.cells[i], valeur, gras=i % 2 == 0, italique=i % 2 == 1)
    saut_de_page(doc)

    titre(doc, "Représentation des personnages et des régions", niveau=3)
    nota_bene(
        doc,
        "seuls les mains characters (personnages principaux) des joueurs/euses ont été "
        "pris en compte pour la création de ce tableau.",
    )
    tableau_tiers(doc, Counter(j["persos"][0] for j in tous if j["persos"]), avec_icones=True)
    doc.add_paragraph()
    tableau_tiers(doc, Counter(j["ville"] for j in tous if j["ville"]), avec_icones=False, regrouper_uniques=True)
    source(doc, f"start.gg – {vierge.TOURNOI}")
    saut_de_page(doc)


def deroule_bracket(doc, joueurs_par_seed, head_to_heads):
    titre(doc, "Déroulé du bracket")
    nota_bene(
        doc,
        "le bracket Singles est définitif, néanmoins, des upsets/DQs peuvent être vecteurs "
        "de changements durant les phases de pools.",
    )
    doc.add_paragraph(
        "Le bracket présenté, ci-dessous, est une projection théorique (dit le Projected), "
        "à l’aide du seeding."
    )
    titre(doc, "Top 32", niveau=2)
    emplacement_image(doc, "Bracket Top 32 – Winners", 8)
    emplacement_image(doc, "Bracket Top 32 – Losers", 6)
    source(doc, f"start.gg – Top 32 – {vierge.TOURNOI}")
    saut_de_page(doc)

    titre(doc, "Seeding", niveau=3)
    table = tableau(doc, ["#", "Top 8", "#", "Top 16", "#", "Top 24", "#", "Top 32"], 8,
                    [0.8, 3.2, 0.8, 3.2, 0.8, 3.2, 0.8, 3.2])
    for seed in range(1, 33):
        joueur = joueurs_par_seed.get(seed)
        cellules = table.rows[1 + (seed - 1) % 8].cells
        colonne = 2 * ((seed - 1) // 8)
        ecrire(cellules[colonne], str(seed), gras=True, alignement=WD_ALIGN_PARAGRAPH.CENTER)
        ecrire(cellules[colonne + 1], joueur["pseudo"] if joueur else "", italique=True)
    source(doc, f"start.gg – {vierge.TOURNOI}")

    titre(doc, "Représentation des personnages", niveau=3)
    top32 = [joueurs_par_seed[s] for s in range(1, 33) if s in joueurs_par_seed]
    tableau_tiers(doc, Counter(p for j in top32 for p in j["persos"]), avec_icones=True)
    source(doc, f"start.gg – Top 32 – {vierge.TOURNOI}")
    saut_de_page(doc)

    titre(doc, "Top 8", niveau=2)
    emplacement_image(doc, "Bracket Top 8", 6)
    source(doc, f"start.gg – Top 8 – {vierge.TOURNOI}")
    titre(doc, "Head-to-Heads et setcounts", niveau=3)
    for joueur_1, joueur_2, phase, resultats in head_to_heads:
        head_to_head(doc, joueur_1, joueur_2, phase, resultats)
    saut_de_page(doc)


def head_to_head(doc, joueur_1, joueur_2, phase, resultats):
    paragraphe = doc.add_paragraph()
    paragraphe.paragraph_format.space_before = Pt(12)
    rouge = RGBColor(0xFF, 0x00, 0x00)
    for texte, couleur, italique in (
        (f"{joueur_1['pseudo']} ", None, False), (f"({joueur_1['seed']})", rouge, False),
        (f" vs {joueur_2['pseudo']} ", None, False), (f"({joueur_2['seed']})", rouge, False),
        (f" – {phase}", None, True),
    ):
        run = paragraphe.add_run(texte)
        run.bold = True
        run.italic = italique
        if couleur:
            run.font.color.rgb = couleur

    victoires = sum(r["score"] > r["score_adversaire"] for r in resultats)
    defaites = sum(r["score"] < r["score_adversaire"] for r in resultats)
    table = tableau(doc, ["", "", "", "", ""], len(resultats), [2.6, 1.4, 7.0, 1.4, 2.6])
    entete = table.rows[0].cells
    entete[0].merge(entete[1]).paragraphs[0].runs[0].text = joueur_1["pseudo"].upper()
    entete[2].paragraphs[0].runs[0].text = f"{victoires} - {defaites}"
    entete[3].merge(entete[4]).paragraphs[0].runs[0].text = joueur_2["pseudo"].upper()
    # Nettoie les paragraphes vides laissés par la fusion
    for cellule in (table.cell(0, 0), table.cell(0, 3)):
        for p in cellule.paragraphs[1:]:
            p._p.getparent().remove(p._p)

    for ligne, r in zip(table.rows[1:], resultats):
        valeurs = (joueur_1["pseudo"], str(r["score"]), f"{r['tournoi']} ({r['round']})",
                   str(r["score_adversaire"]), joueur_2["pseudo"])
        for i, texte in enumerate(valeurs):
            ecrire(ligne.cells[i], texte, gras=True, italique=i == 2, alignement=WD_ALIGN_PARAGRAPH.CENTER)
        gagnant = (0, 1) if r["score"] > r["score_adversaire"] else (3, 4)
        for i in gagnant:
            fond_cellule(ligne.cells[i], VERT)
    if not resultats:
        doc.add_paragraph("Aucune confrontation trouvée sur start.gg.").runs[0].italic = True


# --- Programme principal ----------------------------------------------------

def main():
    parseur = argparse.ArgumentParser(description="Remplit le document de suivi à partir de start.gg.")
    parseur.add_argument("url", nargs="?", help="URL de l'event start.gg (demandée si absente)")
    parseur.add_argument("--sortie", help="nom du fichier .docx généré (demandé si absent)")
    args = parseur.parse_args()
    # Certains terminaux Windows ne savent pas afficher les emojis des noms d'event (🏆…)
    sys.stdout.reconfigure(errors="replace")

    charger_env()
    token = os.environ.get("STARTGG_TOKEN")
    if not token:
        sys.exit("STARTGG_TOKEN manquant : l'ajouter dans le fichier .env (STARTGG_TOKEN=...).")
    api = StartGG(token)

    print("=== Création du document de suivi de tournoi ===")
    evenement = demander_evenement(api, normaliser_slug(args.url) if args.url else None)
    vierge.TOURNOI = evenement["tournament"]["name"]  # utilisé par l'en-tête et la page de garde
    infos = demander_infos()
    sortie_defaut = "Document de suivi - " + re.sub(r'[<>:"/\\|?*]', "", vierge.TOURNOI) + ".docx"
    sortie = args.sortie or demander("Nom du fichier généré", defaut=sortie_defaut)
    if not sortie.lower().endswith(".docx"):
        sortie += ".docx"
    print()

    print("Récupération des pools et des seeds...")
    pools = joueurs_par_pool(api, evenement)
    tous = [j for _, joueurs in pools for j in joueurs]
    harmoniser_villes(tous)
    if not tous:
        sys.exit("Aucun inscrit trouvé sur cet event start.gg.")
    joueurs_par_seed = {j["seed"]: j for j in tous if j["seed"]}

    sans_persos = []
    for i, joueur in enumerate(sorted(tous, key=lambda j: (j["seed"] is None, j["seed"] or 0)), start=1):
        print(f"Personnages {i}/{len(tous)} : {joueur['pseudo']}")
        joueur["persos"] = persos_principaux(api, joueur)
        if not joueur["persos"]:
            sans_persos.append(joueur["pseudo"])

    print("Head-to-heads du Top 8...")
    head_to_heads = []
    for seed_1, seed_2, phase in ((1, 4, "Winners Semi-Final"), (2, 3, "Winners Semi-Final"),
                                  (8, 5, "Top 8 Losers Round 1"), (7, 6, "Top 8 Losers Round 1")):
        joueur_1, joueur_2 = joueurs_par_seed.get(seed_1), joueurs_par_seed.get(seed_2)
        if not (joueur_1 and joueur_2):
            continue
        sets = api.sets_joueur(joueur_1["player_id"], PAGES_MAX_H2H)
        _, resultats = analyser_sets(sets, joueur_1["player_id"], slug_exclu=evenement["slug"])
        resultats = [r for r in resultats if r["adversaire"] == joueur_2["player_id"]]
        head_to_heads.append((joueur_1, joueur_2, phase, resultats))

    doc = Document()
    vierge.configurer_styles(doc)
    vierge.configurer_section(doc.sections[0])
    vierge.page_de_garde(doc, infos["organisation"], infos["participation"], infos["auteur"])
    vierge.sommaire(doc)
    informations_generales(doc, evenement, infos)
    formats_et_ruleset(doc, evenement, infos["format_sets"])
    profils_joueurs(doc, pools, lire_csv("rankings.csv"), lire_csv("teams.csv"))
    deroule_bracket(doc, joueurs_par_seed, head_to_heads)
    vierge.liens_utiles(doc)
    vierge.notes_annexes(doc)

    doc.save(sortie)
    print(f"\nDocument généré : {sortie}")
    if sans_persos:
        print(f"Aucun personnage trouvé sur start.gg pour {len(sans_persos)} joueur(s) :")
        print("  " + ", ".join(sans_persos))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nAnnulé.")
