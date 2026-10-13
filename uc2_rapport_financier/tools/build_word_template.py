"""Construit le modèle Word (contrôles de contenu) à partir du rapport exemple.

Part de la première moitié de « Rapports financiers pages.docx » (exemple 2025
rempli, où les valeurs variables sont surlignées en jaune) et remplace chaque
valeur variable par un contrôle de contenu texte brut. Les contrôles sont lus
par l'action Power Automate « Remplir un modèle Microsoft Word ».

Usage : python tools/build_word_template.py <source.docx> <modele.docx> <controles.json>
"""
import copy
import json
import sys
import zipfile

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# Lignes du tableau « Bilan actif », dans l'ordre du document (ligne 1 à 37).
# Doit rester aligné avec LIGNES dans scripts/BilanActif.ts.
CLES_LIGNES = [
    "ActifsIncorporels", "Placements", "TerrainsConstructions",
    "PlacementsEntreprisesLiees", "AutresPlacements", "CreancesEspecesDeposees",
    "PlacementsUC", "PartCessionnaires", "ProvCotisNonAcquisesNV",
    "ProvAssuranceVie", "ProvPrestationsVie", "ProvPrestationsNV", "ProvPBVie",
    "ProvPBNV", "ProvEgalisationVie", "ProvEgalisationNV", "AutresProvTechVie",
    "AutresProvTechNV", "ProvTechUC", "PartsGarants", "Creances",
    "CreancesOpDirectes", "CreancesReassurance", "AutresCreances", "Personnel",
    "EtatOrganismesSociaux", "DebiteursDivers", "AutresActifs",
    "ActifsCorporelsExploitation", "AvoirsBanque", "ComptesRegularisation",
    "InteretsLoyers", "FraisAcquisitionVie", "FraisAcquisitionNV",
    "AutresComptesRegularisation", "DifferencesConversion", "TotalActif",
]

# Contrôles hors tableau, dans l'ordre d'apparition des surlignages.
CLES_TEXTE = {
    0: ["AnneeN"],                         # « Bilan au 31 décembre 2025 »
    7: ["AnneeN"],                         # « … au 31 décembre 2025. »
    17: ["AnneeN", "PlusValuesLatentes"],  # « Au 31 décembre 2025 … 1 100 451 milliers »
}
INDEX_TABLEAU = 4
DERNIER_INDEX_MODELE = 24  # éléments 0..24 = exemple « Bilan actif » ; la suite est ignorée

ID_DEPART = 1001


def w(tag, **attrs):
    el = etree.Element(W + tag)
    for k, v in attrs.items():
        el.set(W + k, v)
    return el


class Controles:
    def __init__(self):
        self.prochain_id = ID_DEPART
        self.liste = []

    def creer(self, cle, rpr):
        """Contrôle de contenu texte brut au niveau run, affichant [cle]."""
        id_ = self.prochain_id
        self.prochain_id += 1
        self.liste.append({"id": id_, "tag": cle})

        sdt = w("sdt")
        pr = etree.SubElement(sdt, W + "sdtPr")
        if rpr is not None:
            pr.append(copy.deepcopy(rpr))
        pr.append(w("alias", val=cle))
        pr.append(w("tag", val=cle))
        pr.append(w("id", val=str(id_)))
        pr.append(w("text"))
        contenu = etree.SubElement(sdt, W + "sdtContent")
        r = etree.SubElement(contenu, W + "r")
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
        t = etree.SubElement(r, W + "t")
        t.text = "[%s]" % cle
        t.set(XML_SPACE, "preserve")
        return sdt


def rpr_sans_surlignage(run):
    rpr = run.find(W + "rPr")
    if rpr is None:
        return None
    rpr = copy.deepcopy(rpr)
    for tag in ("highlight", "spacing"):
        for el in rpr.findall(W + tag):
            rpr.remove(el)
    return rpr


def est_surligne(run):
    return run.find(f"{W}rPr/{W}highlight") is not None


def remplacer_surlignages(paragraphe, cles, controles):
    """Remplace chaque groupe de runs surlignés consécutifs par un contrôle."""
    runs = [c for c in paragraphe if c.tag == W + "r"]
    groupes, courant = [], []
    for run in runs:
        if est_surligne(run):
            courant.append(run)
        elif courant:
            groupes.append(courant)
            courant = []
    if courant:
        groupes.append(courant)
    if len(groupes) != len(cles):
        raise ValueError(f"{len(groupes)} zones surlignées pour {len(cles)} clés : {cles}")
    for groupe, cle in zip(groupes, cles):
        sdt = controles.creer(cle, rpr_sans_surlignage(groupe[0]))
        paragraphe.replace(groupe[0], sdt)
        for run in groupe[1:]:
            paragraphe.remove(run)


def remplacer_cellule(tc, cle, gras, controles):
    """Vide la cellule de valeur et y place un contrôle aligné à droite."""
    tcpr = tc.find(W + "tcPr")
    if tcpr is not None:
        for vm in tcpr.findall(W + "vMerge"):
            tcpr.remove(vm)
    p = tc.find(W + "p")
    for extra in tc.findall(W + "p")[1:]:
        tc.remove(extra)
    ppr = p.find(W + "pPr")
    for child in list(p):
        if child is not ppr:
            p.remove(child)
    if ppr is None:
        ppr = w("pPr")
        p.insert(0, ppr)
    if ppr.find(W + "jc") is None:
        # Cellule vide dans l'exemple : on reprend la mise en forme des cellules chiffrées
        for el in ppr.findall(W + "rPr"):
            ppr.remove(el)
        ppr.append(w("ind", right="68"))
        ppr.append(w("jc", val="right"))
    rpr = w("rPr")
    if gras:
        rpr.append(w("b"))
    rpr.append(w("color", val="000000"))
    rpr.append(w("sz", val="16"))
    p.append(controles.creer(cle, rpr))


def construire(source, cible, json_controles):
    with zipfile.ZipFile(source) as zin:
        doc = etree.fromstring(zin.read("word/document.xml"))
        body = doc.find(W + "body")
        enfants = list(body)
        controles = Controles()

        for idx, cles in CLES_TEXTE.items():
            remplacer_surlignages(enfants[idx], cles, controles)

        lignes = enfants[INDEX_TABLEAU].findall(W + "tr")
        if len(lignes) != len(CLES_LIGNES) + 1:
            raise ValueError(f"Tableau inattendu : {len(lignes)} lignes")
        entete = lignes[0].findall(W + "tc")
        remplacer_surlignages(entete[1].find(W + "p"), ["AnneeN"], controles)
        remplacer_surlignages(entete[2].find(W + "p"), ["AnneeN1"], controles)
        for tr, cle in zip(lignes[1:], CLES_LIGNES):
            libelle, val_n, val_n1 = tr.findall(W + "tc")
            gras = libelle.find(f".//{W}r/{W}rPr/{W}b") is not None
            remplacer_cellule(val_n, cle + "_N", gras, controles)
            remplacer_cellule(val_n1, cle + "_N1", gras, controles)

        # On garde l'exemple « Bilan actif » (0..24) + les propriétés de section finales
        for el in enfants[DERNIER_INDEX_MODELE + 1:-1]:
            body.remove(el)

        xml = etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)
        with zipfile.ZipFile(cible, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)

    with open(json_controles, "w", encoding="utf-8") as f:
        json.dump(controles.liste, f, ensure_ascii=False, indent=2)
    print(f"{len(controles.liste)} contrôles -> {cible}")


if __name__ == "__main__":
    construire(*sys.argv[1:4])
