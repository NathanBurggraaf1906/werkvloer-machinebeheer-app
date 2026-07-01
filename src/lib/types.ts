export type Afdeling = {
  id: string;
  title: string;
  omschrijving: string;
  actief: boolean;
  volgorde: number;
};

export type Persoon = {
  id: string;
  title: string;
  email: string;
  functie: string;
  afdelingId: string;
  actief: boolean;
};

export type Leverancier = {
  id: string;
  title: string;
  contactpersoon: string;
  email: string;
  telefoon: string;
  opmerking: string;
};

export type Weekdag =
  | "Maandag"
  | "Dinsdag"
  | "Woensdag"
  | "Donderdag"
  | "Vrijdag"
  | "Zaterdag"
  | "Zondag";

export type Machine = {
  id: string;
  title: string;
  afdelingId: string;
  leverancierId: string;
  verantwoordelijkeId: string;
  omschrijving: string;
  status: "Operationeel" | "Onderhoud nodig" | "Buiten gebruik";
  serieNummer: string;
  aankoopdatum: string;
  garantieVerloopdatum: string;
  afbeeldingUrl: string;
  actief: boolean;
};

export type Onderhoud = {
  id: string;
  title: string;
  machineId: string;
  datumGepland: string;
  datumUitgevoerd: string;
  typeOnderhoud: "Preventief" | "Correctief" | "Keuring" | "Schoonmaak";
  verantwoordelijkeId: string;
  verantwoordelijkeType: "Persoon" | "Leverancier";
  verantwoordelijkeRefId: string;
  leverancierId: string;
  status: "Gepland" | "In proces" | "Voltooid";
  herhaling:
    | "Geen"
    | "Wekelijks"
    | "Elke 2 weken"
    | "Maandelijks"
    | "Elk kwartaal"
    | "Elk half jaar"
    | "Jaarlijks";
  herhalingWeekdag: Weekdag;
  herhalingTot: string;
  automatischOpnieuwPlannen: boolean;
  opmerking: string;
};

export type Bijlage = {
  id: string;
  naam: string;
  type: string;
  grootte: number;
  dataUrl: string;
};

export type StoringOpmerking = {
  id: string;
  title: string;
  machineId: string;
  datum: string;
  melderId: string;
  type: "Storing" | "Opmerking" | "Verbeterpunt";
  prioriteit: "Laag" | "Normaal" | "Hoog";
  status: "Open" | "In behandeling" | "Opgelost";
  omschrijving: string;
  oplossing: string;
  bijlagen?: Bijlage[];
};

export type MachineDocument = {
  id: string;
  title: string;
  machineId: string;
  documentType: "Handleiding" | "Keuring" | "Onderhoud" | "Foto" | "Overig";
  omschrijving: string;
  vervaldatum: string;
  actief: boolean;
  url: string;
  bestandNaam?: string;
  bestandType?: string;
  bestandDataUrl?: string;
};

export type AppData = {
  afdelingen: Afdeling[];
  personen: Persoon[];
  leveranciers: Leverancier[];
  machines: Machine[];
  onderhoud: Onderhoud[];
  storingen: StoringOpmerking[];
  documenten: MachineDocument[];
};

export type EntityName = keyof AppData;
