/**
 * Office Script – Bilan actif
 *
 * Lit l'export Excel « Contrôle de report - R1 - Revue du bilan actif »
 * (tableau croisé aplati : Niv1 / Niv2 / Niv3 / compte, une colonne par période)
 * et renvoie, sous forme de chaîne JSON, les valeurs attendues par les
 * contrôles de contenu du modèle Word (word/Modele_Bilan_Actif.docx).
 *
 * - N   = période « AAAA.12 » la plus récente de l'export, N-1 = « (AAAA-1).12 »
 * - Les montants de l'export sont déjà en milliers d'euros : arrondis à l'unité
 * - Les sous-totaux sont la somme des lignes arrondies affichées, pour que le
 *   tableau Word s'additionne
 * - Format français : « 4 333 049 », « -83 », zéro affiché « - »
 *
 * Clés renvoyées : AnneeN, AnneeN1, puis <Cle>_N et <Cle>_N1 pour chaque ligne
 * de LIGNES (mêmes noms que les tags des contrôles Word).
 */

interface Ligne {
  cle: string;
  source?: string[];  // agrégat Excel : [Niv1] ou [Niv1, Niv2] ou [Niv1, Niv2, Niv3]
  enfants?: string[]; // ou somme d'autres lignes du tableau
}

const PART_CESS = "Part des cessionnaires et rétrocessionnaires dans les provisions techniques";
const REGUL = "Comptes de régularisation - Actif";

// Ordre et clés alignés sur CLES_LIGNES dans tools/build_word_template.py
const LIGNES: Ligne[] = [
  { cle: "ActifsIncorporels", source: ["Actifs incorporels"] },
  { cle: "Placements", enfants: ["TerrainsConstructions", "PlacementsEntreprisesLiees", "AutresPlacements", "CreancesEspecesDeposees"] },
  { cle: "TerrainsConstructions", source: ["Placements", "Terrains et constructions"] },
  { cle: "PlacementsEntreprisesLiees", source: ["Placements", "Placements dans les entreprises liées et dans les entreprises avec lesquelles il existe un lien de participation"] },
  { cle: "AutresPlacements", source: ["Placements", "Autres placements"] },
  { cle: "CreancesEspecesDeposees", source: ["Placements", "Créances pour espèces déposées auprès des entreprises cédantes"] },
  { cle: "PlacementsUC", source: ["Placements représentant les provisions techniques afférentes aux opérations en unités de compte"] },
  { cle: "PartCessionnaires", enfants: ["ProvCotisNonAcquisesNV", "ProvAssuranceVie", "ProvPrestationsVie", "ProvPrestationsNV", "ProvPBVie", "ProvPBNV", "ProvEgalisationVie", "ProvEgalisationNV", "AutresProvTechVie", "AutresProvTechNV", "ProvTechUC"] },
  { cle: "ProvCotisNonAcquisesNV", source: [PART_CESS, "Provisions pour cotisations non acquises (Non Vie)"] },
  { cle: "ProvAssuranceVie", source: [PART_CESS, "Provision d'assurance vie"] },
  { cle: "ProvPrestationsVie", source: [PART_CESS, "Provision pour prestations à payer (Vie)"] },
  { cle: "ProvPrestationsNV", source: [PART_CESS, "Provisions pour prestations à payer (Non Vie)"] },
  { cle: "ProvPBVie", source: [PART_CESS, "Provision pour participation aux excédents et ristournes (Vie)"] },
  { cle: "ProvPBNV", source: [PART_CESS, "Provision pour participation aux excédents et ristournes (Non-Vie)"] },
  { cle: "ProvEgalisationVie", source: [PART_CESS, "Provision pour égalisation (Vie)."] },
  { cle: "ProvEgalisationNV", source: [PART_CESS, "Provision pour égalisation (Non Vie)"] },
  { cle: "AutresProvTechVie", source: [PART_CESS, "Autres provisions techniques (Vie)."] },
  { cle: "AutresProvTechNV", source: [PART_CESS, "Autres provisions techniques (Non Vie)"] },
  { cle: "ProvTechUC", source: [PART_CESS, "Provisions techniques des opérations en unité de compte"] },
  { cle: "PartsGarants", source: ["Parts des garants dans les engagements techniques donnés en substitution"] },
  { cle: "Creances", enfants: ["CreancesOpDirectes", "CreancesReassurance", "AutresCreances"] },
  { cle: "CreancesOpDirectes", source: ["Créances", "Créances nées d'opérations directes et de prises en substitution"] },
  { cle: "CreancesReassurance", source: ["Créances", "Créances nées d'opérations de réassurance et de cessions en substitution"] },
  { cle: "AutresCreances", enfants: ["Personnel", "EtatOrganismesSociaux", "DebiteursDivers"] },
  { cle: "Personnel", source: ["Créances", "Autres créances", "Personnel."] },
  { cle: "EtatOrganismesSociaux", source: ["Créances", "Autres créances", "État, organismes sociaux, collectivités publiques"] },
  { cle: "DebiteursDivers", source: ["Créances", "Autres créances", "Débiteurs divers"] },
  { cle: "AutresActifs", enfants: ["ActifsCorporelsExploitation", "AvoirsBanque"] },
  { cle: "ActifsCorporelsExploitation", source: ["Autres actifs", "Actifs corporels d'exploitation"] },
  { cle: "AvoirsBanque", source: ["Autres actifs", "Avoirs en banque, CCP et caisse"] },
  { cle: "ComptesRegularisation", enfants: ["InteretsLoyers", "FraisAcquisitionVie", "FraisAcquisitionNV", "AutresComptesRegularisation"] },
  { cle: "InteretsLoyers", source: [REGUL, "Intérêts et loyers acquis non échus"] },
  { cle: "FraisAcquisitionVie", source: [REGUL, "Frais d'acquisition reportés", "Frais d'acquisition reportés (Vie)"] },
  { cle: "FraisAcquisitionNV", source: [REGUL, "Frais d'acquisition reportés", "Frais d'acquisition reportés (Non-Vie)"] },
  { cle: "AutresComptesRegularisation", source: [REGUL, "Autres comptes de régularisation"] },
  { cle: "DifferencesConversion", source: ["Différence de conversion"] },
  { cle: "TotalActif", enfants: ["ActifsIncorporels", "Placements", "PlacementsUC", "PartCessionnaires", "PartsGarants", "Creances", "AutresActifs", "ComptesRegularisation", "DifferencesConversion"] },
];

function main(workbook: ExcelScript.Workbook): string {
  const feuille = workbook.getWorksheet("source") ?? workbook.getWorksheets()[0];
  const utilisee = feuille.getUsedRange(true);
  if (!utilisee) {
    throw new Error(`La feuille « ${feuille.getName()} » est vide.`);
  }
  // Lecture depuis A1 pour que les colonnes A/B/C (Niv1/2/3) aient les index 0/1/2
  const derniere = utilisee.getLastCell();
  const valeurs = feuille
    .getRangeByIndexes(0, 0, derniere.getRowIndex() + 1, derniere.getColumnIndex() + 1)
    .getValues();

  // 1. Ligne d'en-tête des périodes (« AAAA.MM ») et choix de N / N-1
  const estPeriode = (v: string | number | boolean) => /^\d{4}\.\d{2}$/.test(String(v).trim());
  const iEntete = valeurs.findIndex(ligne => ligne.some(estPeriode));
  if (iEntete < 0) {
    throw new Error("Aucune colonne de période (AAAA.MM) trouvée dans l'export.");
  }
  const colonnes: { [periode: string]: number } = {};
  valeurs[iEntete].forEach((v, j) => {
    if (estPeriode(v)) colonnes[String(v).trim()] = j;
  });
  const anneesDecembre = Object.keys(colonnes)
    .filter(p => p.endsWith(".12"))
    .map(p => Number(p.slice(0, 4)));
  if (anneesDecembre.length === 0) {
    throw new Error("Aucune période de décembre (AAAA.12) dans l'export.");
  }
  const anneeN = Math.max(...anneesDecembre);
  const colN = colonnes[`${anneeN}.12`];
  const colN1 = colonnes[`${anneeN - 1}.12`];
  if (colN1 === undefined) {
    throw new Error(`Période ${anneeN - 1}.12 (N-1) absente de l'export.`);
  }

  // 2. Agrégats par niveau. Les libellés Niv1/2/3 ne sont renseignés que sur la
  //    première ligne de chaque groupe : on les propage aux lignes suivantes.
  const agregats: { [cle: string]: [number, number] } = {};
  const ajouter = (cle: string, n: number, n1: number) => {
    const a = agregats[cle] ?? [0, 0];
    agregats[cle] = [a[0] + n, a[1] + n1];
  };
  const nombre = (v: string | number | boolean) => (typeof v === "number" ? v : 0);
  let niv1 = "", niv2 = "", niv3 = "";
  for (let i = iEntete + 1; i < valeurs.length; i++) {
    const [a, b, c] = valeurs[i].slice(0, 3).map(v => String(v).trim());
    if (normaliser(a) === "total general") break;
    if (normaliser(a).startsWith("total ")) continue;
    if (a) { niv1 = a; niv2 = ""; niv3 = ""; }
    if (b) { niv2 = b; niv3 = ""; }
    if (c) { niv3 = c; }
    const n = nombre(valeurs[i][colN]);
    const n1 = nombre(valeurs[i][colN1]);
    ajouter(cleAgregat([niv1]), n, n1);
    ajouter(cleAgregat([niv1, niv2]), n, n1);
    ajouter(cleAgregat([niv1, niv2, niv3]), n, n1);
  }

  // 3. Valeurs des lignes du tableau Word (arrondies, sous-totaux = somme des enfants)
  const parCle: { [cle: string]: Ligne } = {};
  LIGNES.forEach(l => (parCle[l.cle] = l));
  const calculees: { [cle: string]: [number, number] } = {};
  const manquants: string[] = [];
  const calculer = (cle: string): [number, number] => {
    if (calculees[cle]) return calculees[cle];
    const ligne = parCle[cle];
    let res: [number, number] = [0, 0];
    if (ligne.enfants) {
      ligne.enfants.map(calculer).forEach(([n, n1]) => (res = [res[0] + n, res[1] + n1]));
    } else {
      const agg = agregats[cleAgregat(ligne.source)];
      if (!agg) {
        manquants.push(ligne.source.join(" > "));
      } else {
        res = [arrondir(agg[0]), arrondir(agg[1])];
      }
    }
    calculees[cle] = res;
    return res;
  };
  LIGNES.forEach(l => calculer(l.cle));
  if (manquants.length > 0) {
    throw new Error(`Libellés introuvables dans l'export Excel : ${manquants.join(" | ")}`);
  }

  // 4. Sortie
  const sortie: { [cle: string]: string } = {
    AnneeN: String(anneeN),
    AnneeN1: String(anneeN - 1),
  };
  LIGNES.forEach(l => {
    sortie[`${l.cle}_N`] = formater(calculees[l.cle][0]);
    sortie[`${l.cle}_N1`] = formater(calculees[l.cle][1]);
  });
  return JSON.stringify(sortie);
}

/** Minuscules, sans accents ni ponctuation : « Provision pour égalisation (Vie). » -> « provision pour egalisation vie » */
function normaliser(texte: string): string {
  return texte
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function cleAgregat(libelles: string[]): string {
  return libelles.map(normaliser).join("|");
}

/** Arrondi à l'unité, symétrique pour les négatifs (-83,5 -> -84) */
function arrondir(x: number): number {
  return Math.sign(x) * Math.round(Math.abs(x));
}

/** 4333049 -> « 4 333 049 », -83 -> « -83 », 0 -> « - » */
function formater(n: number): string {
  if (n === 0) return "-";
  const milliers = String(Math.abs(n)).replace(/\B(?=(\d{3})+(?!\d))/g, " ");
  return (n < 0 ? "-" : "") + milliers;
}
