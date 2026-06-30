"use client";

import Image from "next/image";
import {
  AccountInfo,
  InteractionRequiredAuthError,
  PublicClientApplication,
} from "@azure/msal-browser";
import { FormEvent, ReactNode, useEffect, useMemo, useState } from "react";
import { initialData } from "@/lib/seed-data";
import type {
  Afdeling,
  AppData,
  Bijlage,
  EntityName,
  Leverancier,
  Machine,
  MachineDocument,
  Onderhoud,
  Persoon,
  StoringOpmerking,
  Weekdag,
} from "@/lib/types";

const appVersion = "V1.2";
const storageKey = "werkvloer-machinebeheer-v1";
const microsoftClientId = "0d1f2e04-7363-408c-8d69-26516c6f1e98";
const microsoftTenantId = "568c87e9-d6ed-4409-acab-1251c4d47545";
const graphScopes = ["User.Read", "Sites.ReadWrite.All"];
const sharePointHost = "1906makersvancharcuterie.sharepoint.com";
const sharePointSitePath = "/sites/werkvloer-machinebeheer";

type Section = "werkvloer" | "beheer";
type FlowScreen =
  | "start"
  | "afdelingen"
  | "personen"
  | "persoonMachines"
  | "machines"
  | "paspoort"
  | "documenten"
  | "onderhoud"
  | "onderhoudAgenda"
  | "onderhoudDetail"
  | "storingDetail"
  | "storingNieuw"
  | "storingen";
type IconName = "home" | "department" | "person" | "machine" | "document" | "maintenance" | "alert" | "camera" | "plus" | "back" | "spark";
type AuthState =
  | { status: "loading"; msal?: PublicClientApplication }
  | { status: "signedOut"; msal: PublicClientApplication }
  | { status: "signedIn"; account: AccountInfo; msal: PublicClientApplication }
  | { status: "error"; error: string; msal?: PublicClientApplication };
type SharePointListItem = {
  id: string;
  fields?: Record<string, unknown>;
  driveItem?: {
    "@microsoft.graph.downloadUrl"?: string;
    imageUrl?: string;
    name?: string;
    webUrl?: string;
  };
};
type SharePointList = { id: string; displayName?: string; name?: string };
type SharePointColumn = { displayName?: string; name?: string };
type SharePointDrive = { id: string; name?: string };
type SharePointDriveItem = { id: string; name?: string; webUrl?: string };
type MachinePhotoUploadResult = { document: MachineDocument; warning?: string };
type MachineFieldUpdate = Pick<Machine, "status" | "verantwoordelijkeId">;
type SharePointState =
  | { status: "idle"; data?: undefined; error?: undefined }
  | { status: "loading"; data?: undefined; error?: undefined }
  | { status: "ready"; data: AppData; error?: undefined }
  | { status: "error"; data?: undefined; error: string };

const entityLabels: Record<EntityName, string> = {
  afdelingen: "Afdelingen",
  machines: "Machines",
  personen: "Personen",
  leveranciers: "Leveranciers",
  onderhoud: "Onderhoud",
  storingen: "Storingen / opmerkingen",
  documenten: "Machine documenten",
};

const entityOrder: EntityName[] = [
  "afdelingen",
  "machines",
  "personen",
  "leveranciers",
  "onderhoud",
  "storingen",
  "documenten",
];

const weekdagen: Weekdag[] = [
  "Maandag",
  "Dinsdag",
  "Woensdag",
  "Donderdag",
  "Vrijdag",
  "Zaterdag",
  "Zondag",
];

function createMicrosoftClient() {
  return new PublicClientApplication({
    auth: {
      authority: `https://login.microsoftonline.com/${microsoftTenantId}`,
      clientId: microsoftClientId,
      postLogoutRedirectUri: window.location.origin,
      redirectUri: window.location.origin,
    },
    cache: {
      cacheLocation: "localStorage",
    },
  });
}

function createId(prefix: string) {
  return `${prefix}-${Date.now().toString(36)}-${Math.random()
    .toString(36)
    .slice(2, 7)}`;
}

function normalizeData(data: AppData): AppData {
  return {
    ...data,
    onderhoud: data.onderhoud.map((taak) => {
      const legacyStatus = taak.status as string;
      const status: Onderhoud["status"] =
        legacyStatus === "Bezig"
          ? "In proces"
          : legacyStatus === "Uitgevoerd"
            ? "Voltooid"
            : legacyStatus === "Uitgesteld"
              ? "Gepland"
              : (legacyStatus as Onderhoud["status"]);

      return {
        ...taak,
        status,
        herhaling: taak.herhaling ?? "Geen",
        herhalingWeekdag: taak.herhalingWeekdag ?? "Maandag",
        herhalingTot: taak.herhalingTot ?? "",
        verantwoordelijkeType: taak.verantwoordelijkeType ?? "Persoon",
        verantwoordelijkeRefId:
          taak.verantwoordelijkeRefId ?? taak.verantwoordelijkeId ?? "",
      };
    }),
  };
}

function getResponsibleValue(taak: Onderhoud) {
  const type = taak.verantwoordelijkeType ?? "Persoon";
  const id = taak.verantwoordelijkeRefId ?? taak.verantwoordelijkeId ?? "";
  return `${type}:${id}`;
}

function parseResponsibleValue(value: string) {
  const [type, id] = value.split(":");
  return {
    type: type === "Leverancier" ? "Leverancier" : "Persoon",
    id: id ?? "",
  } as const;
}

function getResponsibleLabel(taak: Onderhoud, personen: Persoon[], leveranciers: Leverancier[]) {
  const { type, id } = parseResponsibleValue(getResponsibleValue(taak));
  if (type === "Leverancier") return leveranciers.find((item) => item.id === id)?.title ?? "Leverancier";
  return personen.find((item) => item.id === id)?.title ?? "Persoon";
}

function getWeekdayIndex(weekdag: Weekdag) {
  return [1, 2, 3, 4, 5, 6, 0][weekdagen.indexOf(weekdag)] ?? 1;
}

function getFirstWeekdayDate(year: number, month: number, weekdag: Weekdag) {
  const datum = new Date(year, month, 1);
  const offset = (getWeekdayIndex(weekdag) - datum.getDay() + 7) % 7;
  datum.setDate(1 + offset);
  return datum;
}

function useAppData() {
  const [data, setData] = useState<AppData>(() => {
    if (typeof window === "undefined") {
      return initialData;
    }

    const saved = window.localStorage.getItem(storageKey);
    if (!saved) {
      return initialData;
    }

    try {
      return normalizeData(JSON.parse(saved) as AppData);
    } catch {
      return initialData;
    }
  });

  useEffect(() => {
    window.localStorage.setItem(storageKey, JSON.stringify(data));
  }, [data]);

  function updateEntity<T extends EntityName>(entity: T, records: AppData[T]) {
    setData((current) => ({ ...current, [entity]: records }));
  }

  function resetData() {
    setData(initialData);
    window.localStorage.removeItem(storageKey);
  }

  return { data, updateEntity, resetData };
}

function useMachinePhotoOverrides() {
  const [photos, setPhotos] = useState<Record<string, string>>({});

  function updateMachinePhoto(machineId: string, photoUrl: string) {
    setPhotos((current) => ({ ...current, [machineId]: photoUrl }));
  }

  return { photos, updateMachinePhoto };
}

function useMicrosoftAuth() {
  const [auth, setAuth] = useState<AuthState>({ status: "loading" });

  useEffect(() => {
    let cancelled = false;

    async function initializeAuth() {
      try {
        const msal = createMicrosoftClient();
        await msal.initialize();
        const redirectResult = await msal.handleRedirectPromise();
        const account =
          redirectResult?.account ?? msal.getActiveAccount() ?? msal.getAllAccounts()[0];

        if (cancelled) return;

        if (account) {
          msal.setActiveAccount(account);
          setAuth({ account, msal, status: "signedIn" });
        } else {
          setAuth({ msal, status: "signedOut" });
        }
      } catch (error) {
        if (!cancelled) {
          setAuth({
            error: error instanceof Error ? error.message : "Aanmelden is niet gelukt.",
            status: "error",
          });
        }
      }
    }

    initializeAuth();

    return () => {
      cancelled = true;
    };
  }, []);

  async function signIn() {
    if (!("msal" in auth) || !auth.msal) return;
    await auth.msal.loginRedirect({
      scopes: graphScopes,
    });
  }

  async function signOut() {
    if (!("msal" in auth) || !auth.msal) return;
    await auth.msal.logoutRedirect({
      account: auth.status === "signedIn" ? auth.account : undefined,
      postLogoutRedirectUri: window.location.origin,
    });
  }

  return { auth, signIn, signOut };
}

async function graphFetch<T>(path: string, accessToken: string): Promise<T> {
  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Graph ${response.status}: ${body || response.statusText}`);
  }

  return response.json() as Promise<T>;
}

async function graphFetchRaw(path: string, accessToken: string, init: RequestInit) {
  const headers = new Headers(init.headers);
  headers.set("Authorization", `Bearer ${accessToken}`);

  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...init,
    headers,
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Graph ${response.status}: ${body || response.statusText}`);
  }

  if (response.status === 204) return undefined;
  return response.json() as Promise<unknown>;
}

async function graphFetchBlob(path: string, accessToken: string) {
  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Graph ${response.status}: ${body || response.statusText}`);
  }

  return response.blob();
}

function blobToDataUrl(blob: Blob) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

async function getSharePointToken(msal: PublicClientApplication, account: AccountInfo) {
  try {
    const token = await msal.acquireTokenSilent({
      account,
      scopes: graphScopes,
    });

    return token.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      await msal.acquireTokenRedirect({
        account,
        scopes: graphScopes,
      });
    }

    throw error;
  }
}

function stringValue(fields: Record<string, unknown>, names: string[], fallback = "") {
  for (const name of names) {
    const value = fieldValue(fields, name);
    if (typeof value === "string") return value;
    if (typeof value === "number") return String(value);
  }

  return fallback;
}

function booleanValue(fields: Record<string, unknown>, names: string[], fallback = false) {
  for (const name of names) {
    const value = fieldValue(fields, name);
    if (typeof value === "boolean") return value;
    if (typeof value === "number") return value === 1;
    if (typeof value === "string") return value === "1" || value.toLowerCase() === "true" || value.toLowerCase() === "yes";
  }

  return fallback;
}

function numberValue(fields: Record<string, unknown>, names: string[], fallback = 0) {
  for (const name of names) {
    const value = fieldValue(fields, name);
    if (typeof value === "number") return value;
    if (typeof value === "string" && value.trim()) return Number(value);
  }

  return fallback;
}

function dateValue(fields: Record<string, unknown>, names: string[]) {
  const raw = stringValue(fields, names);
  return raw ? raw.slice(0, 10) : "";
}

function normalizeFieldName(name: string) {
  return name
    .replace(/_x0020_/gi, "")
    .replace(/[^a-z0-9]/gi, "")
    .toLowerCase();
}

function fieldValue(fields: Record<string, unknown>, name: string) {
  if (name in fields) return fields[name];

  const normalizedName = normalizeFieldName(name);
  const matchingKey = Object.keys(fields).find((key) => normalizeFieldName(key) === normalizedName);
  return matchingKey ? fields[matchingKey] : undefined;
}

function lookupIdValue(fields: Record<string, unknown>, displayName: string) {
  const compact = displayName.replace(/\s+/g, "");
  const candidates = [
    `${displayName}LookupId`,
    `${compact}LookupId`,
    `${displayName}Id`,
    `${compact}Id`,
  ];

  for (const candidate of candidates) {
    const value = fieldValue(fields, candidate);
    if (typeof value === "number") return String(value);
    if (typeof value === "string" && value.trim()) return value;
  }

  return "";
}

function choiceValue<T extends string>(value: string, allowed: readonly T[], fallback: T) {
  return allowed.includes(value as T) ? (value as T) : fallback;
}

function idFromLookup(map: Map<string, string>, lookupId: string) {
  return lookupId ? map.get(lookupId) ?? "" : "";
}

function idFromLookupTitle(map: Map<string, string>, lookupTitle: string) {
  return lookupTitle ? map.get(normalizeFieldName(lookupTitle)) ?? "" : "";
}

function sharePointItemIdFromAppId(id: string) {
  return id.split("-").at(-1) ?? "";
}

function safeSharePointFileName(name: string) {
  return name
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[<>:"/\\|?*#%{}~&]/g, "-")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 80);
}

function fileExtension(file: File) {
  const extension = file.name.split(".").pop();
  if (extension && extension !== file.name) return extension.toLowerCase();
  if (file.type === "image/png") return "png";
  if (file.type === "image/webp") return "webp";
  return "jpg";
}

function columnInternalName(columns: SharePointColumn[], displayName: string) {
  const normalized = normalizeFieldName(displayName);
  return columns.find((column) => normalizeFieldName(column.displayName ?? column.name ?? "") === normalized)?.name;
}

async function fetchListItems(
  siteId: string,
  lists: SharePointList[],
  listName: string,
  accessToken: string,
  includeDriveItem = false,
) {
  const list = lists.find((item) => item.displayName === listName || item.name === listName);
  if (!list) return [];

  const result = await graphFetch<{ value: SharePointListItem[] }>(
    `/sites/${siteId}/lists/${list.id}/items?$expand=${includeDriveItem ? "fields,driveItem" : "fields"}&$top=200`,
    accessToken,
  );

  if (includeDriveItem) {
    await Promise.all(
      result.value.map(async (item) => {
        if (!item.driveItem) return;
        const documentType = choiceValue(
          stringValue(item.fields ?? {}, ["Documenttype", "Document type"]),
          ["Handleiding", "Keuring", "Onderhoud", "Foto", "Overig"],
          "Overig",
        );

        if (documentType === "Foto") {
          try {
            const blob = await graphFetchBlob(
              `/sites/${siteId}/lists/${list.id}/items/${item.id}/driveItem/content`,
              accessToken,
            );
            item.driveItem.imageUrl = await blobToDataUrl(blob);
            return;
          } catch {
            // Fall back to Graph thumbnails below.
          }
        }

        try {
          const thumbnail = await graphFetch<{
            large?: { url?: string };
            medium?: { url?: string };
            small?: { url?: string };
          }>(
            `/sites/${siteId}/lists/${list.id}/items/${item.id}/driveItem/thumbnails/0`,
            accessToken,
          );
          item.driveItem.imageUrl = thumbnail.large?.url ?? thumbnail.medium?.url ?? thumbnail.small?.url;
        } catch {
          item.driveItem.imageUrl = item.driveItem["@microsoft.graph.downloadUrl"];
        }
      }),
    );
  }

  return result.value;
}

function mapSharePointData(items: {
  afdelingen: SharePointListItem[];
  personen: SharePointListItem[];
  leveranciers: SharePointListItem[];
  machines: SharePointListItem[];
  onderhoud: SharePointListItem[];
  storingen: SharePointListItem[];
  documenten: SharePointListItem[];
}): AppData {
  const afdelingIds = new Map<string, string>();
  const persoonIds = new Map<string, string>();
  const leverancierIds = new Map<string, string>();
  const machineIds = new Map<string, string>();

  const afdelingen = items.afdelingen.map((item): Afdeling => {
    const fields = item.fields ?? {};
    const id = `sp-afdeling-${item.id}`;
    afdelingIds.set(item.id, id);

    return {
      actief: booleanValue(fields, ["Actief"], true),
      id,
      omschrijving: stringValue(fields, ["Omschrijving"]),
      title: stringValue(fields, ["Title", "Titel"], "Afdeling"),
      volgorde: numberValue(fields, ["Volgorde"], 0),
    };
  });

  const personen = items.personen.map((item): Persoon => {
    const fields = item.fields ?? {};
    const id = `sp-persoon-${item.id}`;
    persoonIds.set(item.id, id);

    return {
      actief: booleanValue(fields, ["Actief"], true),
      afdelingId: idFromLookup(afdelingIds, lookupIdValue(fields, "Afdeling")),
      email: stringValue(fields, ["Email"]),
      functie: stringValue(fields, ["Functie"]),
      id,
      title: stringValue(fields, ["Title", "Titel"], "Persoon"),
    };
  });

  const leveranciers = items.leveranciers.map((item): Leverancier => {
    const fields = item.fields ?? {};
    const id = `sp-leverancier-${item.id}`;
    leverancierIds.set(item.id, id);

    return {
      contactpersoon: stringValue(fields, ["Contactpersoon"]),
      email: stringValue(fields, ["Email"]),
      id,
      opmerking: stringValue(fields, ["Opmerking"]),
      telefoon: stringValue(fields, ["Telefoon"]),
      title: stringValue(fields, ["Title", "Titel"], "Leverancier"),
    };
  });

  const machines = items.machines.map((item): Machine => {
    const fields = item.fields ?? {};
    const id = `sp-machine-${item.id}`;
    machineIds.set(item.id, id);

    return {
      aankoopdatum: dateValue(fields, ["Aankoopdatum"]),
      actief: booleanValue(fields, ["Actief"], true),
      afbeeldingUrl: stringValue(fields, ["AfbeeldingURL", "Afbeelding URL"]),
      afdelingId: idFromLookup(afdelingIds, lookupIdValue(fields, "Afdeling")),
      garantieVerloopdatum: dateValue(fields, ["Garantieverloopdatum", "Garantie verloopdatum"]),
      id,
      leverancierId: idFromLookup(leverancierIds, lookupIdValue(fields, "Leverancier")),
      omschrijving: stringValue(fields, ["Omschrijving"]),
      serieNummer: stringValue(fields, ["Serienummer", "Serie nummer"]),
      status: choiceValue(stringValue(fields, ["Status"]), ["Operationeel", "Onderhoud nodig", "Buiten gebruik"], "Operationeel"),
      title: stringValue(fields, ["Title", "Titel"], "Machine"),
      verantwoordelijkeId: idFromLookup(persoonIds, lookupIdValue(fields, "Verantwoordelijke")),
    };
  });
  const machineTitleIds = new Map(machines.map((machine) => [normalizeFieldName(machine.title), machine.id]));

  const onderhoud = items.onderhoud.map((item): Onderhoud => {
    const fields = item.fields ?? {};
    const verantwoordelijkePersoonId = idFromLookup(persoonIds, lookupIdValue(fields, "Verantwoordelijke persoon"));
    const verantwoordelijkeLeverancierId = idFromLookup(leverancierIds, lookupIdValue(fields, "Verantwoordelijke leverancier"));
    const verantwoordelijkeType = verantwoordelijkeLeverancierId ? "Leverancier" : "Persoon";
    const verantwoordelijkeRefId = verantwoordelijkeLeverancierId || verantwoordelijkePersoonId;

    return {
      datumGepland: dateValue(fields, ["Datumgepland", "Datum gepland"]),
      datumUitgevoerd: dateValue(fields, ["Datumuitgevoerd", "Datum uitgevoerd"]),
      herhaling: choiceValue(stringValue(fields, ["Herhaling"]), ["Geen", "Wekelijks", "Maandelijks", "Jaarlijks", "Eerste weekdag van de maand", "Eerste weekdag van het kwartaal"], "Geen"),
      herhalingTot: dateValue(fields, ["Herhalentot", "Herhalen tot"]),
      herhalingWeekdag: choiceValue(stringValue(fields, ["Herhalingweekdag", "Herhaling weekdag"]), weekdagen, "Maandag"),
      id: `sp-onderhoud-${item.id}`,
      leverancierId: verantwoordelijkeLeverancierId,
      machineId: idFromLookup(machineIds, lookupIdValue(fields, "Machine")),
      opmerking: stringValue(fields, ["Opmerking"]),
      status: choiceValue(stringValue(fields, ["Status"]), ["Gepland", "In proces", "Voltooid"], "Gepland"),
      title: stringValue(fields, ["Title", "Titel"], "Onderhoud"),
      typeOnderhoud: choiceValue(stringValue(fields, ["Typeonderhoud", "Type onderhoud"]), ["Preventief", "Correctief", "Keuring", "Schoonmaak"], "Preventief"),
      verantwoordelijkeId: verantwoordelijkePersoonId,
      verantwoordelijkeRefId,
      verantwoordelijkeType,
    };
  });

  const storingen = items.storingen.map((item): StoringOpmerking => {
    const fields = item.fields ?? {};

    return {
      datum: dateValue(fields, ["Datum"]),
      id: `sp-storing-${item.id}`,
      machineId: idFromLookup(machineIds, lookupIdValue(fields, "Machine")),
      melderId: idFromLookup(persoonIds, lookupIdValue(fields, "Melder")),
      omschrijving: stringValue(fields, ["Omschrijving"]),
      oplossing: stringValue(fields, ["Oplossing"]),
      prioriteit: choiceValue(stringValue(fields, ["Prioriteit"]), ["Laag", "Normaal", "Hoog"], "Normaal"),
      status: choiceValue(stringValue(fields, ["Status"]), ["Open", "In behandeling", "Opgelost"], "Open"),
      title: stringValue(fields, ["Title", "Titel"], "Melding"),
      type: choiceValue(stringValue(fields, ["Meldingtype", "Melding type"]), ["Storing", "Opmerking", "Verbeterpunt"], "Storing"),
    };
  });

  const documenten = items.documenten.map((item): MachineDocument => {
    const fields = item.fields ?? {};
    const documentType = choiceValue(stringValue(fields, ["Documenttype", "Document type"]), ["Handleiding", "Keuring", "Onderhoud", "Foto", "Overig"], "Overig");
    const sharePointUrl = item.driveItem?.webUrl ?? stringValue(fields, ["FileRef", "Link"]);
    const imageUrl = item.driveItem?.imageUrl ?? item.driveItem?.["@microsoft.graph.downloadUrl"];
    const url = documentType === "Foto" ? imageUrl ?? sharePointUrl : sharePointUrl;

    return {
      actief: booleanValue(fields, ["Actief"], true),
      documentType,
      id: `sp-document-${item.id}`,
      machineId:
        idFromLookup(machineIds, lookupIdValue(fields, "Machine")) ||
        idFromLookupTitle(machineTitleIds, stringValue(fields, ["Machine"])),
      omschrijving: stringValue(fields, ["Omschrijving"]),
      title: item.driveItem?.name ?? stringValue(fields, ["FileLeafRef", "Title", "Titel"], "Document"),
      url,
      vervaldatum: dateValue(fields, ["Vervaldatum"]),
    };
  });

  return normalizeData({ afdelingen, documenten, leveranciers, machines, onderhoud, personen, storingen });
}

function useSharePointData(auth: AuthState) {
  const [state, setState] = useState<SharePointState>({ status: "idle" });

  useEffect(() => {
    if (auth.status !== "signedIn") {
      return;
    }

    const { account, msal } = auth;
    let cancelled = false;

    async function loadSharePointData() {
      setState({ status: "loading" });

      try {
        const accessToken = await getSharePointToken(msal, account);
        const site = await graphFetch<{ id: string }>(
          `/sites/${sharePointHost}:${sharePointSitePath}`,
          accessToken,
        );
        const lists = await graphFetch<{ value: SharePointList[] }>(
          `/sites/${site.id}/lists?$select=id,displayName,name`,
          accessToken,
        );

        const [
          afdelingen,
          personen,
          leveranciers,
          machines,
          onderhoud,
          storingen,
          documenten,
        ] = await Promise.all([
          fetchListItems(site.id, lists.value, "Afdelingen", accessToken),
          fetchListItems(site.id, lists.value, "Personen", accessToken),
          fetchListItems(site.id, lists.value, "Leveranciers", accessToken),
          fetchListItems(site.id, lists.value, "Machines", accessToken),
          fetchListItems(site.id, lists.value, "Onderhoud", accessToken),
          fetchListItems(site.id, lists.value, "Storingen Opmerkingen", accessToken),
          fetchListItems(site.id, lists.value, "Machine documenten", accessToken, true),
        ]);

        if (!cancelled) {
          setState({
            data: mapSharePointData({
              afdelingen,
              documenten,
              leveranciers,
              machines,
              onderhoud,
              personen,
              storingen,
            }),
            status: "ready",
          });
        }
      } catch (error) {
        if (!cancelled) {
          setState({
            error: error instanceof Error ? error.message : "SharePoint-data laden is niet gelukt.",
            status: "error",
          });
        }
      }
    }

    loadSharePointData();

    return () => {
      cancelled = true;
    };
  }, [auth]);

  return state;
}

async function uploadMachinePhotoToSharePoint(
  auth: AuthState,
  machine: Machine,
  file: File,
): Promise<MachinePhotoUploadResult> {
  if (auth.status !== "signedIn") {
    throw new Error("Je moet ingelogd zijn om een foto naar SharePoint te uploaden.");
  }

  const accessToken = await getSharePointToken(auth.msal, auth.account);
  const site = await graphFetch<{ id: string }>(
    `/sites/${sharePointHost}:${sharePointSitePath}`,
    accessToken,
  );
  const [lists, drives] = await Promise.all([
    graphFetch<{ value: SharePointList[] }>(
      `/sites/${site.id}/lists?$select=id,displayName,name`,
      accessToken,
    ),
    graphFetch<{ value: SharePointDrive[] }>(
      `/sites/${site.id}/drives?$select=id,name`,
      accessToken,
    ),
  ]);

  const documentLibrary = lists.value.find(
    (list) => list.displayName === "Machine documenten" || list.name === "Machine documenten",
  );
  const drive = drives.value.find((item) => item.name === "Machine documenten");

  if (!documentLibrary || !drive) {
    throw new Error("Documentbibliotheek 'Machine documenten' is niet gevonden in SharePoint.");
  }

  const extension = fileExtension(file);
  const fileName = `${safeSharePointFileName(machine.title || "machine")}-foto-${Date.now()}.${extension}`;
  const uploaded = await graphFetchRaw(
    `/sites/${site.id}/drives/${drive.id}/root:/${encodeURIComponent(fileName)}:/content`,
    accessToken,
    {
      body: await file.arrayBuffer(),
      headers: {
        "Content-Type": file.type || "application/octet-stream",
      },
      method: "PUT",
    },
  ) as SharePointDriveItem;

  let warning: string | undefined;

  try {
    const columns = await graphFetch<{ value: SharePointColumn[] }>(
      `/sites/${site.id}/lists/${documentLibrary.id}/columns?$select=name,displayName`,
      accessToken,
    );
    const fields: Record<string, string | boolean> = {};
    const machineColumn = columnInternalName(columns.value, "Machine");
    const documentTypeColumn = columnInternalName(columns.value, "Document type");
    const omschrijvingColumn = columnInternalName(columns.value, "Omschrijving");
    const actiefColumn = columnInternalName(columns.value, "Actief");
    const machineItemId = sharePointItemIdFromAppId(machine.id);

    if (machineColumn && machineItemId) fields[`${machineColumn}LookupId`] = machineItemId;
    if (documentTypeColumn) fields[documentTypeColumn] = "Foto";
    if (omschrijvingColumn) fields[omschrijvingColumn] = `Machinefoto voor ${machine.title}`;
    if (actiefColumn) fields[actiefColumn] = true;

    if (Object.keys(fields).length > 0) {
      await graphFetchRaw(
        `/sites/${site.id}/drives/${drive.id}/items/${uploaded.id}/listItem/fields`,
        accessToken,
        {
          body: JSON.stringify(fields),
          headers: {
            "Content-Type": "application/json",
          },
          method: "PATCH",
        },
      );
    }
  } catch (error) {
    warning = error instanceof Error
      ? `Foto is geupload, maar metadata koppelen lukte nog niet: ${error.message}`
      : "Foto is geupload, maar metadata koppelen lukte nog niet.";
  }

  return {
    document: {
      actief: true,
      bestandNaam: uploaded.name ?? fileName,
      bestandType: file.type || "image/jpeg",
      documentType: "Foto",
      id: createId("sp-photo"),
      machineId: machine.id,
      omschrijving: `Machinefoto voor ${machine.title}`,
      title: uploaded.name ?? fileName,
      url: uploaded.webUrl ?? "",
      vervaldatum: "",
    },
    warning,
  };
}

async function updateMachineFieldsInSharePoint(
  auth: AuthState,
  machine: Machine,
  update: MachineFieldUpdate,
) {
  if (auth.status !== "signedIn") {
    throw new Error("Je moet ingelogd zijn om machinegegevens op te slaan.");
  }

  const machineItemId = sharePointItemIdFromAppId(machine.id);
  const verantwoordelijkeItemId = sharePointItemIdFromAppId(update.verantwoordelijkeId);

  if (!machineItemId) {
    throw new Error("Deze machine heeft geen SharePoint-id.");
  }

  const accessToken = await getSharePointToken(auth.msal, auth.account);
  const site = await graphFetch<{ id: string }>(
    `/sites/${sharePointHost}:${sharePointSitePath}`,
    accessToken,
  );
  const lists = await graphFetch<{ value: SharePointList[] }>(
    `/sites/${site.id}/lists?$select=id,displayName,name`,
    accessToken,
  );
  const machinesList = lists.value.find(
    (list) => list.displayName === "Machines" || list.name === "Machines",
  );

  if (!machinesList) {
    throw new Error("SharePoint-lijst 'Machines' is niet gevonden.");
  }

  const columns = await graphFetch<{ value: SharePointColumn[] }>(
    `/sites/${site.id}/lists/${machinesList.id}/columns?$select=name,displayName`,
    accessToken,
  );
  const fields: Record<string, string> = {};
  const statusColumn = columnInternalName(columns.value, "Status");
  const verantwoordelijkeColumn = columnInternalName(columns.value, "Verantwoordelijke");

  if (statusColumn) fields[statusColumn] = update.status;
  if (verantwoordelijkeColumn && verantwoordelijkeItemId) {
    fields[`${verantwoordelijkeColumn}LookupId`] = verantwoordelijkeItemId;
  }

  if (Object.keys(fields).length === 0) {
    throw new Error("Geen passende SharePoint-kolommen gevonden voor Status/Verantwoordelijke.");
  }

  await graphFetchRaw(
    `/sites/${site.id}/lists/${machinesList.id}/items/${machineItemId}/fields`,
    accessToken,
    {
      body: JSON.stringify(fields),
      headers: {
        "Content-Type": "application/json",
      },
      method: "PATCH",
    },
  );
}

async function fileToBijlage(file: File): Promise<Bijlage> {
  const dataUrl = await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });

  return {
    id: createId("bij"),
    naam: file.name,
    type: file.type || "application/octet-stream",
    grootte: file.size,
    dataUrl,
  };
}

export default function Home() {
  const { auth, signIn, signOut } = useMicrosoftAuth();
  const { data: localData, updateEntity } = useAppData();
  const { photos: machinePhotos, updateMachinePhoto } = useMachinePhotoOverrides();
  const sharePointData = useSharePointData(auth);
  const [machineFieldOverrides, setMachineFieldOverrides] = useState<Record<string, Partial<Machine>>>({});
  const [onderhoudOverrides, setOnderhoudOverrides] = useState<Onderhoud[] | null>(null);
  const [storingenOverrides, setStoringenOverrides] = useState<StoringOpmerking[] | null>(null);
  const sourceData = sharePointData.status === "ready" ? sharePointData.data : localData;
  const data = useMemo(
    () => ({
      ...sourceData,
      onderhoud: onderhoudOverrides ?? sourceData.onderhoud,
      storingen: storingenOverrides ?? sourceData.storingen,
      machines: sourceData.machines.map((machine) => {
        const documentPhoto = sourceData.documenten.find(
          (document) => document.machineId === machine.id && document.documentType === "Foto" && document.actief && document.url,
        );

        return {
          ...machine,
          ...machineFieldOverrides[machine.id],
          afbeeldingUrl: documentPhoto?.url ?? machine.afbeeldingUrl ?? machinePhotos[machine.id] ?? "",
        };
      }),
    }),
    [machineFieldOverrides, machinePhotos, onderhoudOverrides, sourceData, storingenOverrides],
  );
  const [section, setSection] = useState<Section>("werkvloer");
  const [flowScreen, setFlowScreen] = useState<FlowScreen>("start");
  const [selectedAfdelingId, setSelectedAfdelingId] = useState("");
  const [selectedPersoonId, setSelectedPersoonId] = useState("");
  const [selectedMachineId, setSelectedMachineId] = useState("");
  const [activeEntity, setActiveEntity] = useState<EntityName>("machines");

  const activeAfdelingen = useMemo(
    () =>
      [...data.afdelingen]
        .filter((afdeling) => afdeling.actief)
        .sort((a, b) => a.volgorde - b.volgorde),
    [data.afdelingen],
  );

  const selectedAfdeling = data.afdelingen.find(
    (afdeling) => afdeling.id === selectedAfdelingId,
  );
  const selectedPersoon = data.personen.find(
    (persoon) => persoon.id === selectedPersoonId,
  );
  const filteredMachines = data.machines.filter(
    (machine) => machine.actief && machine.afdelingId === selectedAfdelingId,
  );
  const selectedMachine = data.machines.find(
    (machine) => machine.id === selectedMachineId,
  );

  function goHome() {
    setSection("werkvloer");
    setFlowScreen("start");
    setSelectedAfdelingId("");
    setSelectedPersoonId("");
    setSelectedMachineId("");
  }

  async function uploadMachinePhoto(machine: Machine, file: File) {
    const result = await uploadMachinePhotoToSharePoint(auth, machine, file);
    const previewUrl = URL.createObjectURL(file);
    updateMachinePhoto(machine.id, previewUrl);
    updateEntity("documenten", [{ ...result.document, url: previewUrl }, ...data.documenten]);
    return result;
  }

  async function updateMachinePassport(machine: Machine, update: MachineFieldUpdate) {
    await updateMachineFieldsInSharePoint(auth, machine, update);
    setMachineFieldOverrides((current) => ({
      ...current,
      [machine.id]: {
        ...current[machine.id],
        ...update,
      },
    }));
  }

  function updateOnderhoudRecords(onderhoud: Onderhoud[]) {
    setOnderhoudOverrides(onderhoud);
    updateEntity("onderhoud", onderhoud);
  }

  function updateStoringenRecords(storingen: StoringOpmerking[]) {
    setStoringenOverrides(storingen);
    updateEntity("storingen", storingen);
  }

  if (auth.status === "loading") {
    return <AuthShell message="Microsoft-login laden..." />;
  }

  if (auth.status === "signedOut") {
    return <LoginScreen onSignIn={signIn} />;
  }

  if (auth.status === "error") {
    return <LoginScreen error={auth.error} onSignIn={signIn} />;
  }

  return (
    <main className="appShell">
      <header className="topbar">
        <div className="brandZone">
          <button className="brandButton" onClick={goHome} type="button" aria-label="Terug naar start">
            <Image
              alt="1906 makers van charcuterie"
              className="brandLogo"
              height={86}
              priority
              src="/brand/1906-round-logo.jpg"
              width={86}
            />
          </button>
          <button className="homeInBrand" onClick={goHome} type="button">
            <LineIcon name="home" />
            Home
          </button>
        </div>
        <div className="topActions" aria-label="Hoofdnavigatie">
          <UserBadge account={auth.account} onSignOut={signOut} />
          <div className="navSwitch">
            <button
              className={section === "werkvloer" ? "smallButton topNavButton active" : "smallButton topNavButton"}
              onClick={goHome}
              type="button"
            >
              Werkvloer
            </button>
            <button
              className={section === "beheer" ? "smallButton topNavButton active" : "smallButton topNavButton"}
              onClick={() => setSection("beheer")}
              type="button"
            >
              Beheer
            </button>
          </div>
        </div>
      </header>

      <SharePointNotice state={sharePointData} />

      {sharePointData.status !== "ready" && sharePointData.status !== "error" ? (
        <SharePointLoadingScreen />
      ) : section === "werkvloer" ? (
        <WerkvloerFlow
          afdelingen={activeAfdelingen}
          data={data}
          filteredMachines={filteredMachines}
          flowScreen={flowScreen}
          goHome={goHome}
          selectedAfdeling={selectedAfdeling}
          selectedMachine={selectedMachine}
          selectedPersoon={selectedPersoon}
          setFlowScreen={setFlowScreen}
          setSelectedAfdelingId={setSelectedAfdelingId}
          setSelectedMachineId={setSelectedMachineId}
          setSelectedPersoonId={setSelectedPersoonId}
          updateMachinePassport={updateMachinePassport}
          updateDocumenten={(documenten) => updateEntity("documenten", documenten)}
          uploadMachinePhoto={uploadMachinePhoto}
          updateOnderhoud={updateOnderhoudRecords}
          updateStoringen={updateStoringenRecords}
        />
      ) : (
        <BeheerView
          activeEntity={activeEntity}
          data={data}
          setActiveEntity={setActiveEntity}
          updateEntity={updateEntity}
        />
      )}
    </main>
  );
}

function AuthShell({ message }: { message: string }) {
  return (
    <main className="appShell authShell">
      <section className="loginCard">
        <div className="loginContent">
          <Image
            alt="1906 makers van charcuterie"
            className="loginLogo"
            height={104}
            priority
            src="/brand/1906-round-logo.jpg"
            width={104}
          />
          <p className="eyebrow">Werkvloer Machinebeheer</p>
          <h1>{message}</h1>

          <span className="loginVersion">Appversie {appVersion}</span>
        </div>
        <Image
          alt="Biologische achterham van 1906 makers van charcuterie"
          className="loginProductPhoto"
          height={360}
          priority
          src="/brand/1906-achterham-hero.jpg"
          width={560}
        />
      </section>
    </main>
  );
}

function SharePointLoadingScreen() {
  return (
    <section className="mobileScreen loadingScreen">
      <div className="loadingCard">
        <LineIcon name="document" />
        <p className="eyebrow">SharePoint</p>
        <h1>Data laden...</h1>
        <p>Afdelingen, personen, machines, onderhoud en storingen worden uit SharePoint opgehaald.</p>
      </div>
    </section>
  );
}

function LoginScreen({ error, onSignIn }: { error?: string; onSignIn: () => void }) {
  return (
    <main className="appShell authShell">
      <section className="loginCard">
        <div className="loginContent">
          <Image
            alt="1906 makers van charcuterie"
            className="loginLogo"
            height={112}
            priority
            src="/brand/1906-round-logo.jpg"
            width={112}
          />
          <p className="eyebrow">Beveiligde werkvloer-app</p>
          <h1>Log in met Microsoft.</h1>
          <p>
            Alleen gebruikers uit de Microsoft 365-omgeving van 1906 makers van
            charcuterie kunnen deze pilot openen.
          </p>
          {error && <p className="authError">{error}</p>}
          <button className="submitButton red" onClick={onSignIn} type="button">
            <LineIcon name="person" />
            Aanmelden met Microsoft
          </button>
          <span className="loginVersion">Appversie {appVersion}</span>
        </div>
        <Image
          alt="Grillworst van 1906 makers van charcuterie"
          className="loginProductPhoto"
          height={360}
          priority
          src="/brand/1906-grillworst-hero.jpg"
          width={560}
        />
      </section>
    </main>
  );
}

function UserBadge({ account, onSignOut }: { account: AccountInfo; onSignOut: () => void }) {
  const displayName = account.name || account.username;

  return (
    <div className="userBadge">
      <span>
        <small>Ingelogd</small>
        <strong>{displayName}</strong>
      </span>
      <button className="logoutButton" onClick={onSignOut} type="button">
        Afmelden
      </button>
    </div>
  );
}

function WerkvloerFlow({
  data,
  afdelingen,
  filteredMachines,
  flowScreen,
  goHome,
  selectedAfdeling,
  selectedMachine,
  selectedPersoon,
  setFlowScreen,
  setSelectedAfdelingId,
  setSelectedMachineId,
  setSelectedPersoonId,
  updateMachinePassport,
  updateDocumenten,
  uploadMachinePhoto,
  updateOnderhoud,
  updateStoringen,
}: {
  data: AppData;
  afdelingen: Afdeling[];
  filteredMachines: Machine[];
  flowScreen: FlowScreen;
  goHome: () => void;
  selectedAfdeling?: Afdeling;
  selectedMachine?: Machine;
  selectedPersoon?: Persoon;
  setFlowScreen: (screen: FlowScreen) => void;
  setSelectedAfdelingId: (id: string) => void;
  setSelectedMachineId: (id: string) => void;
  setSelectedPersoonId: (id: string) => void;
  updateMachinePassport: (machine: Machine, update: MachineFieldUpdate) => Promise<void>;
  updateDocumenten: (documenten: MachineDocument[]) => void;
  uploadMachinePhoto: (machine: Machine, file: File) => Promise<MachinePhotoUploadResult>;
  updateOnderhoud: (onderhoud: Onderhoud[]) => void;
  updateStoringen: (storingen: StoringOpmerking[]) => void;
}) {
  const [selectedDocumentId, setSelectedDocumentId] = useState("");
  const [editingOnderhoudId, setEditingOnderhoudId] = useState("");
  const [editingStoringId, setEditingStoringId] = useState("");
  const [maintenanceFilter, setMaintenanceFilter] = useState("all");
  const [maintenanceStatusFilter, setMaintenanceStatusFilter] = useState("all");
  const [saveNotice, setSaveNotice] = useState("");
  const machineDocuments = data.documenten.filter(
    (document) => document.machineId === selectedMachine?.id && document.actief,
  );
  const machineOnderhoud = data.onderhoud.filter(
    (onderhoud) => onderhoud.machineId === selectedMachine?.id,
  );
  const machineStoringen = data.storingen.filter(
    (storing) => storing.machineId === selectedMachine?.id,
  );
  const verantwoordelijke = data.personen.find(
    (persoon) => persoon.id === selectedMachine?.verantwoordelijkeId,
  );
  const leverancier = data.leveranciers.find(
    (item) => item.id === selectedMachine?.leverancierId,
  );
  const activePersonen = data.personen.filter((persoon) => persoon.actief);
  const personMachines = data.machines.filter(
    (machine) => machine.actief && machine.verantwoordelijkeId === selectedPersoon?.id,
  );
  const selectedDocument = data.documenten.find(
    (document) => document.id === selectedDocumentId,
  );
  const editingOnderhoud = data.onderhoud.find(
    (taak) => taak.id === editingOnderhoudId,
  );
  const editingStoring = data.storingen.find(
    (storing) => storing.id === editingStoringId,
  );

  function chooseAfdeling(afdelingId: string) {
    setSelectedAfdelingId(afdelingId);
    setSelectedPersoonId("");
    setSelectedMachineId("");
    setFlowScreen("machines");
  }

  function choosePersoon(persoonId: string) {
    setSelectedPersoonId(persoonId);
    setSelectedAfdelingId("");
    setSelectedMachineId("");
    setFlowScreen("persoonMachines");
  }

  function chooseMachine(machineId: string) {
    setSelectedMachineId(machineId);
    const machine = data.machines.find((item) => item.id === machineId);
    if (machine) setSelectedAfdelingId(machine.afdelingId);
    setSelectedDocumentId("");
    setEditingOnderhoudId("");
    setEditingStoringId("");
    setFlowScreen("paspoort");
  }

  async function addDocument(formData: FormData) {
    if (!selectedMachine) return;

    const file = formData.get("bestand");
    const bijlage = file instanceof File && file.size > 0 ? await fileToBijlage(file) : undefined;
    const document: MachineDocument = {
      id: createId("doc"),
      title: String(formData.get("title") || bijlage?.naam || "Nieuw document"),
      machineId: selectedMachine.id,
      documentType: String(formData.get("documentType") || "Overig") as MachineDocument["documentType"],
      omschrijving: String(formData.get("omschrijving") || ""),
      vervaldatum: String(formData.get("vervaldatum") || ""),
      actief: true,
      url: bijlage?.dataUrl || "#",
      bestandNaam: bijlage?.naam,
      bestandType: bijlage?.type,
      bestandDataUrl: bijlage?.dataUrl,
    };

    updateDocumenten([document, ...data.documenten]);
    setSelectedDocumentId(document.id);
  }

  async function addMachinePhoto(formData: FormData) {
    if (!selectedMachine) return;

    const file = formData.get("machineFoto");
    if (!(file instanceof File) || file.size === 0) return;

    try {
      const result = await uploadMachinePhoto(selectedMachine, file);
      setSaveNotice(result.warning ?? "Machinefoto opgeslagen in SharePoint.");
    } catch (error) {
      setSaveNotice(error instanceof Error ? error.message : "Machinefoto uploaden naar SharePoint is niet gelukt.");
    }
  }

  async function editMachinePassport(formData: FormData) {
    if (!selectedMachine) return;

    const update: MachineFieldUpdate = {
      status: String(formData.get("status") || selectedMachine.status) as Machine["status"],
      verantwoordelijkeId: String(formData.get("verantwoordelijkeId") || selectedMachine.verantwoordelijkeId),
    };

    try {
      await updateMachinePassport(selectedMachine, update);
      setSaveNotice("Machinepaspoort opgeslagen in SharePoint.");
    } catch (error) {
      setSaveNotice(error instanceof Error ? error.message : "Machinepaspoort opslaan is niet gelukt.");
    }
  }

  function editDocument(documentId: string, formData: FormData) {
    const bijgewerkt = data.documenten.map((document) =>
      document.id === documentId
        ? {
            ...document,
            title: String(formData.get("title") || document.title),
            documentType: String(formData.get("documentType") || document.documentType) as MachineDocument["documentType"],
            omschrijving: String(formData.get("omschrijving") || ""),
            vervaldatum: String(formData.get("vervaldatum") || ""),
            actief: formData.get("actief") === "on",
          }
        : document,
    );

    updateDocumenten(bijgewerkt);
  }

  function addOnderhoud(formData: FormData) {
    if (!selectedMachine) return;
    const responsible = parseResponsibleValue(String(formData.get("verantwoordelijke") || `Persoon:${selectedMachine.verantwoordelijkeId || data.personen[0]?.id || ""}`));

    const taak: Onderhoud = {
      id: createId("ond"),
      title: String(formData.get("title") || "Nieuwe onderhoudstaak"),
      machineId: selectedMachine.id,
      datumGepland: String(formData.get("datumGepland") || ""),
      datumUitgevoerd: "",
      typeOnderhoud: String(formData.get("typeOnderhoud") || "Preventief") as Onderhoud["typeOnderhoud"],
      verantwoordelijkeId: responsible.type === "Persoon" ? responsible.id : "",
      verantwoordelijkeType: responsible.type,
      verantwoordelijkeRefId: responsible.id,
      leverancierId: responsible.type === "Leverancier" ? responsible.id : selectedMachine.leverancierId,
      status: String(formData.get("status") || "Gepland") as Onderhoud["status"],
      herhaling: String(formData.get("herhaling") || "Geen") as Onderhoud["herhaling"],
      herhalingWeekdag: String(formData.get("herhalingWeekdag") || "Maandag") as Weekdag,
      herhalingTot: "",
      opmerking: String(formData.get("opmerking") || ""),
    };

    updateOnderhoud([taak, ...data.onderhoud]);
    setEditingOnderhoudId("");
    setFlowScreen("onderhoud");
  }

  function editOnderhoud(taakId: string, formData: FormData) {
    const responsible = parseResponsibleValue(String(formData.get("verantwoordelijke") || ""));
    const bijgewerkt = data.onderhoud.map((taak) =>
      taak.id === taakId
        ? {
            ...taak,
            title: String(formData.get("title") || taak.title),
            datumGepland: String(formData.get("datumGepland") || ""),
            datumUitgevoerd: String(formData.get("datumUitgevoerd") || ""),
            typeOnderhoud: String(formData.get("typeOnderhoud") || taak.typeOnderhoud) as Onderhoud["typeOnderhoud"],
            verantwoordelijkeId: responsible.type === "Persoon" ? responsible.id : "",
            verantwoordelijkeType: responsible.type,
            verantwoordelijkeRefId: responsible.id,
            leverancierId: responsible.type === "Leverancier" ? responsible.id : taak.leverancierId,
            status: String(formData.get("status") || taak.status) as Onderhoud["status"],
            herhaling: String(formData.get("herhaling") || "Geen") as Onderhoud["herhaling"],
            herhalingWeekdag: String(formData.get("herhalingWeekdag") || "Maandag") as Weekdag,
            herhalingTot: "",
            opmerking: String(formData.get("opmerking") || ""),
          }
        : taak,
    );

    updateOnderhoud(bijgewerkt);
    setEditingOnderhoudId("");
    setFlowScreen("onderhoud");
  }

  function createFollowUpOnderhoud(taak: Onderhoud, datumUitgevoerd?: string) {
    const afgerondeTaak = datumUitgevoerd ? { ...taak, datumUitgevoerd } : taak;
    const nextDate = getNextOnderhoudDate(afgerondeTaak);
    if (taak.herhaling === "Geen" || (taak.herhalingTot && nextDate > taak.herhalingTot)) return null;

    return {
      ...taak,
      datumGepland: nextDate,
      datumUitgevoerd: "",
      id: createId("ond"),
      status: "Gepland" as Onderhoud["status"],
      title: taak.title.replace(/( - vervolg)+$/u, ""),
    };
  }

  function updateOnderhoudStatus(taakId: string, status: Onderhoud["status"]) {
    const vandaag = getTodayValue();
    const taakVoorUpdate = data.onderhoud.find((taak) => taak.id === taakId);
    const wordtNetVoltooid = status === "Voltooid" && taakVoorUpdate?.status !== "Voltooid";
    const datumUitgevoerd = status === "Voltooid" ? taakVoorUpdate?.datumUitgevoerd || vandaag : "";
    const vervolgTaak = wordtNetVoltooid && taakVoorUpdate ? createFollowUpOnderhoud(taakVoorUpdate, datumUitgevoerd) : null;
    const bijgewerkt = data.onderhoud.map((taak) =>
      taak.id === taakId
        ? {
            ...taak,
            datumUitgevoerd,
            status,
          }
        : taak,
    );

    updateOnderhoud(vervolgTaak ? [vervolgTaak, ...bijgewerkt] : bijgewerkt);
  }

  function completeOnderhoud(taakId: string) {
    const taak = data.onderhoud.find((item) => item.id === taakId);
    updateOnderhoudStatus(taakId, taak?.status === "Voltooid" ? "Gepland" : "Voltooid");
  }

  function completeOnderhoudAndReturn(taakId: string) {
    updateOnderhoudStatus(taakId, "Voltooid");
    setEditingOnderhoudId("");
    setFlowScreen("onderhoud");
  }

  function deleteOnderhoud(taakId: string) {
    const taak = data.onderhoud.find((item) => item.id === taakId);
    if (!taak) return;
    const confirmed = window.confirm(`Onderhoudstaak "${taak.title}" verwijderen?`);
    if (!confirmed) return;

    updateOnderhoud(data.onderhoud.filter((item) => item.id !== taakId));
    setEditingOnderhoudId("");
    setFlowScreen("onderhoud");
  }

  function rescheduleOnderhoud(taak: Onderhoud) {
    if (!selectedMachine) return;

    const nieuweTaak = createFollowUpOnderhoud({ ...taak, machineId: selectedMachine.id });
    if (!nieuweTaak) return;

    updateOnderhoud([nieuweTaak, ...data.onderhoud]);
    setEditingOnderhoudId(nieuweTaak.id);
    setFlowScreen("onderhoudDetail");
  }

  function openOnderhoudDetail(taakId: string) {
    setEditingOnderhoudId(taakId);
    setFlowScreen("onderhoudDetail");
  }

  function openNewOnderhoud() {
    setEditingOnderhoudId("");
    setFlowScreen("onderhoudDetail");
  }

  async function addStoring(formData: FormData) {
    if (!selectedMachine) return;

    const files = formData.getAll("bijlagen").filter((file): file is File => file instanceof File && file.size > 0);
    const bijlagen = await Promise.all(files.map(fileToBijlage));
    const melding: StoringOpmerking = {
      id: createId("sto"),
      title: String(formData.get("title") || "Nieuwe melding"),
      machineId: selectedMachine.id,
      datum: new Date().toISOString().slice(0, 10),
      melderId: data.personen[0]?.id ?? "",
      type: String(formData.get("type") || "Opmerking") as StoringOpmerking["type"],
      prioriteit: String(formData.get("prioriteit") || "Normaal") as StoringOpmerking["prioriteit"],
      status: "Open",
      omschrijving: String(formData.get("omschrijving") || ""),
      oplossing: "",
      bijlagen,
    };

    updateStoringen([melding, ...data.storingen]);
    setSaveNotice("Melding opgeslagen op dit apparaat. In de volgende fase koppelen we dit aan gedeelde SharePoint-data.");
    setEditingStoringId(melding.id);
    setFlowScreen("storingen");
  }

  async function editStoring(storingId: string, formData: FormData) {
    const files = formData.getAll("bijlagen").filter((file): file is File => file instanceof File && file.size > 0);
    const nieuweBijlagen = await Promise.all(files.map(fileToBijlage));
    const bijgewerkt = data.storingen.map((storing) =>
      storing.id === storingId
        ? {
            ...storing,
            bijlagen: [...(storing.bijlagen ?? []), ...nieuweBijlagen],
            omschrijving: String(formData.get("omschrijving") || ""),
            oplossing: String(formData.get("oplossing") || ""),
            prioriteit: String(formData.get("prioriteit") || storing.prioriteit) as StoringOpmerking["prioriteit"],
            status: String(formData.get("status") || storing.status) as StoringOpmerking["status"],
            title: String(formData.get("title") || storing.title),
            type: String(formData.get("type") || storing.type) as StoringOpmerking["type"],
          }
        : storing,
    );

    updateStoringen(bijgewerkt);
    setSaveNotice("Storing bijgewerkt.");
    setFlowScreen("storingen");
  }

  function resolveStoring(storingId: string) {
    const vandaag = getTodayValue();
    const bijgewerkt = data.storingen.map((storing) =>
      storing.id === storingId
        ? {
            ...storing,
            oplossing: storing.oplossing || `Opgelost op ${vandaag}`,
            status: "Opgelost" as StoringOpmerking["status"],
          }
        : storing,
    );

    updateStoringen(bijgewerkt);
    setSaveNotice(`Storing opgelost op ${vandaag}.`);
  }

  function openStoringDetail(storingId: string) {
    setEditingStoringId(storingId);
    setFlowScreen("storingDetail");
  }

  if (flowScreen === "start") {
    return (
      <section className="mobileScreen heroScreen">
        <PilotNotice />
        <div className="heroIntro">
          <div className="heroCopy">
            <p className="eyebrow">Mobiele werkvloer-app</p>
            <h1>Machinebeheer zonder omwegen.</h1>
            <p>
              Kies een afdeling, open een machinepaspoort en registreer snel wat er
              speelt op de werkvloer.
            </p>
          </div>
          <Image
            alt="Biologische achterham van 1906 makers van charcuterie"
            className="heroProductPhoto"
            height={360}
            priority
            src="/brand/1906-achterham-hero.jpg"
            width={560}
          />
        </div>
        <div className="heroActions">
          <ActionButton icon="department" label="Afdeling kiezen" onClick={() => setFlowScreen("afdelingen")} />
          <ActionButton icon="person" label="Persoon kiezen" onClick={() => setFlowScreen("personen")} tone="gold" />
          <ActionButton icon="alert" label="Laatste meldingen" onClick={() => setFlowScreen("storingen")} tone="light" />
        </div>
      </section>
    );
  }

  if (flowScreen === "afdelingen") {
    return (
      <section className="mobileScreen">
        <ScreenHeader eyebrow="Stap 1" goHome={goHome} onBack={() => setFlowScreen("start")} title="Afdeling kiezen" />
        <div className="cardList">
          {afdelingen.map((afdeling) => (
            <button className="navCard" key={afdeling.id} onClick={() => chooseAfdeling(afdeling.id)} type="button">
              <LineIcon name="department" />
              <span>
                <strong>{afdeling.title}</strong>
                <small>{afdeling.omschrijving}</small>
              </span>
            </button>
          ))}
        </div>
      </section>
    );
  }

  if (flowScreen === "personen") {
    return (
      <section className="mobileScreen">
        <ScreenHeader eyebrow="Stap 1" goHome={goHome} onBack={() => setFlowScreen("start")} title="Persoon kiezen" />
        <div className="cardList">
          {activePersonen.map((persoon) => (
            <button className="navCard" key={persoon.id} onClick={() => choosePersoon(persoon.id)} type="button">
              <LineIcon name="person" />
              <span>
                <strong>{persoon.title}</strong>
                <small>{persoon.functie} - {data.afdelingen.find((afdeling) => afdeling.id === persoon.afdelingId)?.title ?? "Geen afdeling"}</small>
              </span>
            </button>
          ))}
        </div>
      </section>
    );
  }

  if (flowScreen === "persoonMachines") {
    return (
      <section className="mobileScreen">
        <ScreenHeader eyebrow={selectedPersoon?.title ?? "Persoon"} goHome={goHome} onBack={() => setFlowScreen("personen")} title="Machines per persoon" />
        <div className="cardList">
          {personMachines.map((machine) => (
            <button className="navCard" key={machine.id} onClick={() => chooseMachine(machine.id)} type="button">
              <LineIcon name="machine" />
              <span>
                <strong>{machine.title}</strong>
                <small>{machine.status} - {data.afdelingen.find((afdeling) => afdeling.id === machine.afdelingId)?.title ?? "Geen afdeling"}</small>
              </span>
            </button>
          ))}
          {personMachines.length === 0 && <p className="emptyState">Geen machines gekoppeld aan deze persoon.</p>}
        </div>
      </section>
    );
  }

  if (flowScreen === "machines") {
    return (
      <section className="mobileScreen">
        <ScreenHeader eyebrow={selectedAfdeling?.title ?? "Afdeling"} goHome={goHome} onBack={() => setFlowScreen("afdelingen")} title="Machines" />
        <div className="cardList">
          {filteredMachines.map((machine) => (
            <button className="navCard" key={machine.id} onClick={() => chooseMachine(machine.id)} type="button">
              <LineIcon name="machine" />
              <span>
                <strong>{machine.title}</strong>
                <small>{machine.status} - {machine.serieNummer}</small>
              </span>
            </button>
          ))}
          {filteredMachines.length === 0 && <p className="emptyState">Geen actieve machines in deze afdeling.</p>}
        </div>
      </section>
    );
  }

  if (!selectedMachine && flowScreen !== "storingen") {
    return (
      <section className="mobileScreen">
        <ScreenHeader eyebrow="Geen machine gekozen" goHome={goHome} onBack={() => setFlowScreen("machines")} title="Machinepaspoort" />
        <p className="emptyState">Kies eerst een machine.</p>
      </section>
    );
  }

  if (flowScreen === "paspoort" && selectedMachine) {
    return (
      <section className="mobileScreen">
        <ScreenHeader eyebrow={selectedAfdeling?.title ?? "Machine"} goHome={goHome} onBack={() => setFlowScreen(selectedPersoon ? "persoonMachines" : "machines")} title={selectedMachine.title} />
        {saveNotice && <p className="saveNotice">{saveNotice}</p>}
        <div className="passportCard">
          <MachinePhotoPanel machine={selectedMachine} onSubmit={addMachinePhoto} />
          <section className="passportSection">
            <h2 className="sectionTitleBar"><LineIcon name="person" /> Status & verantwoordelijke</h2>
            <span className={`statusPill ${selectedMachine.status === "Operationeel" ? "good" : "alert"}`}>{selectedMachine.status}</span>
            <p>{selectedMachine.omschrijving}</p>
            <MachinePassportEditForm
              machine={selectedMachine}
              onSubmit={editMachinePassport}
              personen={activePersonen}
            />
          </section>
          <section className="passportSection">
            <h2 className="sectionTitleBar"><LineIcon name="machine" /> Machinegegevens</h2>
            <dl className="passportMeta">
              <div><dt>Serie</dt><dd>{selectedMachine.serieNummer}</dd></div>
              <div><dt>Leverancier</dt><dd>{leverancier?.title ?? "-"}</dd></div>
              <div><dt>Garantie</dt><dd>{selectedMachine.garantieVerloopdatum || "-"}</dd></div>
            </dl>
          </section>
        </div>
        <section className="passportSection actionSection">
          <h2 className="sectionTitleBar"><LineIcon name="spark" /> Acties</h2>
          <div className="actionGrid">
          <ActionButton icon="document" label="Documenten" onClick={() => setFlowScreen("documenten")} />
          <ActionButton icon="maintenance" label="Onderhoud" onClick={() => setFlowScreen("onderhoud")} tone="gold" />
          <ActionButton icon="alert" label="Storingen" onClick={() => setFlowScreen("storingen")} tone="red" />
          </div>
        </section>
        <MachineAiTool
          documents={machineDocuments}
          machine={selectedMachine}
          onderhoud={machineOnderhoud}
          storingen={machineStoringen}
        />
      </section>
    );
  }

  if (flowScreen === "documenten") {
    return (
      <section className="mobileScreen">
        <ScreenHeader eyebrow={selectedMachine?.title ?? "Machine"} goHome={goHome} onBack={() => setFlowScreen("paspoort")} title="Documenten" />
        {selectedDocument ? (
          <DocumentDetail
            document={selectedDocument}
            onBack={() => setSelectedDocumentId("")}
            onSubmit={(formData) => editDocument(selectedDocument.id, formData)}
          />
        ) : (
          <>
            {selectedMachine && <DocumentUploadForm onSubmit={addDocument} />}
            <DocumentList documenten={machineDocuments} onOpen={setSelectedDocumentId} />
          </>
        )}
      </section>
    );
  }

  if (flowScreen === "onderhoud") {
    return (
      <section className="mobileScreen">
        <ScreenHeader eyebrow={selectedMachine?.title ?? "Machine"} goHome={goHome} onBack={() => setFlowScreen("paspoort")} title="Onderhoud" />
        {selectedMachine && (
          <OnderhoudPanel
            machine={selectedMachine}
            onderhoud={machineOnderhoud}
            personen={data.personen}
            leveranciers={data.leveranciers}
            verantwoordelijke={verantwoordelijke}
            onComplete={completeOnderhoud}
            onCreate={openNewOnderhoud}
            onOpenAgenda={() => setFlowScreen("onderhoudAgenda")}
            onOpenTask={openOnderhoudDetail}
          />
        )}
      </section>
    );
  }

  if (flowScreen === "onderhoudAgenda") {
    return (
      <section className="mobileScreen">
        <ScreenHeader eyebrow={selectedMachine?.title ?? "Machine"} goHome={goHome} onBack={() => setFlowScreen("onderhoud")} title="Onderhoud agenda" />
        {selectedMachine && (
          <OnderhoudAgendaView
            filterValue={maintenanceFilter}
            leveranciers={data.leveranciers}
            onderhoud={machineOnderhoud}
            onComplete={completeOnderhoud}
            onCreate={openNewOnderhoud}
            onFilterChange={setMaintenanceFilter}
            onOpenTask={openOnderhoudDetail}
            onStatusFilterChange={setMaintenanceStatusFilter}
            personen={data.personen}
            statusFilter={maintenanceStatusFilter}
          />
        )}
      </section>
    );
  }

  if (flowScreen === "onderhoudDetail") {
    return (
      <section className="mobileScreen">
        <ScreenHeader eyebrow={selectedMachine?.title ?? "Machine"} goHome={goHome} onBack={() => setFlowScreen("onderhoud")} title={editingOnderhoud ? "Onderhoudstaak" : "Nieuwe onderhoudstaak"} />
        {selectedMachine && (
          <OnderhoudDetailView
            defaultResponsibleId={selectedMachine.verantwoordelijkeId}
            leveranciers={data.leveranciers}
            onCancel={() => setFlowScreen("onderhoud")}
            onCreate={addOnderhoud}
            onEdit={editOnderhoud}
            onReschedule={rescheduleOnderhoud}
            onCompleteAndReturn={completeOnderhoudAndReturn}
            onDelete={deleteOnderhoud}
            personen={data.personen}
            taak={editingOnderhoud?.machineId === selectedMachine.id ? editingOnderhoud : undefined}
          />
        )}
      </section>
    );
  }

  if (flowScreen === "storingNieuw") {
    return (
      <section className="mobileScreen">
        <ScreenHeader
          eyebrow={selectedMachine?.title ?? "Machine"}
          goHome={goHome}
          onBack={() => setFlowScreen("storingen")}
          title="Nieuwe storing"
        />
        {saveNotice && <p className="saveNotice">{saveNotice}</p>}
        {selectedMachine ? (
          <StoringForm mode="create" onSubmit={addStoring} />
        ) : (
          <p className="emptyState">Kies eerst een machine voordat je een storing toevoegt.</p>
        )}
      </section>
    );
  }

  if (flowScreen === "storingDetail") {
    return (
      <section className="mobileScreen">
        <ScreenHeader
          eyebrow={selectedMachine?.title ?? "Machine"}
          goHome={goHome}
          onBack={() => setFlowScreen("storingen")}
          title="Storing bekijken"
        />
        {saveNotice && <p className="saveNotice">{saveNotice}</p>}
        {editingStoring ? (
          <StoringForm
            mode="edit"
            onSubmit={(formData) => editStoring(editingStoring.id, formData)}
            storing={editingStoring}
          />
        ) : (
          <p className="emptyState">Deze storing is niet gevonden.</p>
        )}
      </section>
    );
  }

  return (
    <section className="mobileScreen">
      <ScreenHeader
        eyebrow={selectedMachine?.title ?? "Alle machines"}
        goHome={goHome}
        onBack={() => (selectedMachine ? setFlowScreen("paspoort") : setFlowScreen("start"))}
        title="Storingen / opmerkingen"
      />
      {saveNotice && <p className="saveNotice">{saveNotice}</p>}
      {selectedMachine && (
        <section className="maintenanceActions">
          <button className="submitButton red" onClick={() => setFlowScreen("storingNieuw")} type="button">
            <LineIcon name="plus" />
            Nieuwe storing
          </button>
        </section>
      )}
      <StoringList
        onOpen={openStoringDetail}
        onResolve={selectedMachine ? resolveStoring : undefined}
        storingen={selectedMachine ? machineStoringen : data.storingen}
      />
    </section>
  );
}

function SharePointNotice({ state }: { state: SharePointState }) {
  if (state.status === "idle") return null;

  if (state.status === "loading") {
    return (
      <p className="dataNotice loading">
        SharePoint-data wordt geladen...
      </p>
    );
  }

  if (state.status === "error") {
    return (
      <p className="dataNotice error">
        SharePoint-data kon niet geladen worden. De app toont nu lokale testdata. Fout: {state.error}
      </p>
    );
  }

  return (
    <p className="dataNotice ready">
      <strong>SharePoint actief</strong>
      <span>Data wordt gelezen uit SharePoint.</span>
    </p>
  );
}

function PilotNotice() {
  return (
    <p className="pilotNotice">
      <strong>Appversie {appVersion}</strong>
      <span>Lezen uit SharePoint is actief. Nieuwe meldingen en wijzigingen worden in deze stap nog per apparaat/browser bewaard.</span>
    </p>
  );
}

function ScreenHeader({ eyebrow, goHome, onBack, title }: { eyebrow: string; goHome: () => void; onBack: () => void; title: string }) {
  return (
    <div className="screenHeader">
      <button className="smallButton" onClick={onBack} type="button"><LineIcon name="back" /> Terug</button>
      <div>
        <p className="eyebrow">{eyebrow}</p>
        <h1>{title}</h1>
      </div>
      <button className="smallButton homeSmall" onClick={goHome} type="button"><LineIcon name="home" /> Home</button>
    </div>
  );
}

function ActionButton({ icon, label, onClick, tone = "dark" }: { icon: IconName; label: string; onClick: () => void; tone?: "dark" | "gold" | "red" | "light" }) {
  return (
    <button className={`actionButton ${tone}`} onClick={onClick} type="button">
      <LineIcon name={icon} />
      {label}
    </button>
  );
}

function MachinePhotoPanel({ machine, onSubmit }: { machine: Machine; onSubmit: (formData: FormData) => Promise<void> }) {
  const [busy, setBusy] = useState(false);

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const form = event.currentTarget;
    setBusy(true);
    await onSubmit(new FormData(form));
    form.reset();
    setBusy(false);
  }

  return (
    <section className="machinePhotoPanel">
      <h2 className="sectionTitleBar"><LineIcon name="camera" /> Machinefoto</h2>
      {machine.afbeeldingUrl ? (
        <img
          alt={`Foto van ${machine.title}`}
          className="machinePhoto"
          src={machine.afbeeldingUrl}
        />
      ) : (
        <div className="machinePhotoPlaceholder">
          <LineIcon name="camera" />
          <span>Nog geen machinefoto</span>
        </div>
      )}
      <form className="machinePhotoForm" onSubmit={handleSubmit}>
        <label className="fileBox">
          <LineIcon name="camera" />
          Foto toevoegen
          <input accept="image/*" capture="environment" name="machineFoto" type="file" />
        </label>
        <button className="smallButton" disabled={busy} type="submit">
          {busy ? "Opslaan..." : "Foto opslaan"}
        </button>
      </form>
    </section>
  );
}

function MachinePassportEditForm({
  machine,
  onSubmit,
  personen,
}: {
  machine: Machine;
  onSubmit: (formData: FormData) => Promise<void>;
  personen: Persoon[];
}) {
  const [busy, setBusy] = useState(false);

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const form = event.currentTarget;
    setBusy(true);
    await onSubmit(new FormData(form));
    setBusy(false);
  }

  return (
    <form className="passportEditForm" onSubmit={handleSubmit}>
      <label>
        <span>Status</span>
        <select name="status" defaultValue={machine.status}>
          <option>Operationeel</option>
          <option>Onderhoud nodig</option>
          <option>Buiten gebruik</option>
        </select>
      </label>
      <label>
        <span>Verantwoordelijke</span>
        <select name="verantwoordelijkeId" defaultValue={machine.verantwoordelijkeId}>
          {personen.map((persoon) => (
            <option key={persoon.id} value={persoon.id}>
              {persoon.title}
            </option>
          ))}
        </select>
      </label>
      <button className="smallButton" disabled={busy} type="submit">
        {busy ? "Opslaan..." : "Paspoort opslaan"}
      </button>
    </form>
  );
}

function DocumentUploadForm({ onSubmit }: { onSubmit: (formData: FormData) => Promise<void> }) {
  const [busy, setBusy] = useState(false);

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const form = event.currentTarget;
    setBusy(true);
    await onSubmit(new FormData(form));
    form.reset();
    setBusy(false);
  }

  return (
    <form className="entryForm" onSubmit={handleSubmit}>
      <h2><LineIcon name="document" /> Document uploaden</h2>
      <label>Titel<input name="title" placeholder="Bijvoorbeeld: Handleiding" /></label>
      <label>Document type<select name="documentType" defaultValue="Handleiding"><option>Handleiding</option><option>Keuring</option><option>Onderhoud</option><option>Foto</option><option>Overig</option></select></label>
      <label>Omschrijving<textarea name="omschrijving" rows={2} placeholder="Korte toelichting" /></label>
      <label>Vervaldatum<input name="vervaldatum" type="date" /></label>
      <label className="fileBox"><LineIcon name="document" /> Bestand kiezen<input name="bestand" type="file" /></label>
      <button className="submitButton" disabled={busy} type="submit">{busy ? "Uploaden..." : "Document opslaan"}</button>
    </form>
  );
}

function getTodayValue() {
  return new Date().toISOString().slice(0, 10);
}

function isOnderhoudAchterstallig(taak: Onderhoud) {
  return Boolean(taak.datumGepland && taak.datumGepland < getTodayValue() && taak.status !== "Voltooid");
}

function getOnderhoudDisplayStatus(taak: Onderhoud) {
  return isOnderhoudAchterstallig(taak) ? "Achterstallig" : taak.status;
}

function getNextOnderhoudDate(taak: Onderhoud) {
  const basisDatum =
    taak.datumUitgevoerd && taak.datumUitgevoerd > (taak.datumGepland || "")
      ? taak.datumUitgevoerd
      : taak.datumGepland || taak.datumUitgevoerd || getTodayValue();
  const start = new Date(`${basisDatum}T00:00:00`);

  if (taak.herhaling === "Wekelijks") start.setDate(start.getDate() + 7);
  if (taak.herhaling === "Maandelijks") start.setMonth(start.getMonth() + 1);
  if (taak.herhaling === "Jaarlijks") start.setFullYear(start.getFullYear() + 1);
  if (taak.herhaling === "Eerste weekdag van de maand") {
    start.setMonth(start.getMonth() + 1);
    return getFirstWeekdayDate(start.getFullYear(), start.getMonth(), taak.herhalingWeekdag).toISOString().slice(0, 10);
  }
  if (taak.herhaling === "Eerste weekdag van het kwartaal") {
    start.setMonth(start.getMonth() + 3);
    const kwartaalMaand = Math.floor(start.getMonth() / 3) * 3;
    return getFirstWeekdayDate(start.getFullYear(), kwartaalMaand, taak.herhalingWeekdag).toISOString().slice(0, 10);
  }
  if (taak.herhaling === "Geen") start.setDate(start.getDate() + 7);

  return start.toISOString().slice(0, 10);
}

function OnderhoudPanel({
  machine,
  onderhoud,
  personen,
  leveranciers,
  verantwoordelijke,
  onComplete,
  onCreate,
  onOpenAgenda,
  onOpenTask,
}: {
  machine: Machine;
  onderhoud: Onderhoud[];
  personen: Persoon[];
  leveranciers: Leverancier[];
  verantwoordelijke?: Persoon;
  onComplete: (taakId: string) => void;
  onCreate: () => void;
  onOpenAgenda: () => void;
  onOpenTask: (taakId: string) => void;
}) {
  const sortedOnderhoud = [...onderhoud].sort((a, b) => (a.datumGepland || "9999").localeCompare(b.datumGepland || "9999"));

  return (
    <>
      <section className="maintenanceHero">
        <p className="eyebrow">Machine</p>
        <h2>{machine.title}</h2>
        <p className="responsibleLine">Verantwoordelijke persoon: <strong>{verantwoordelijke?.title ?? "-"}</strong></p>
      </section>
      <MaintenanceStats onderhoud={onderhoud} />
      <section className="maintenanceActions">
        <button className="primaryButton" onClick={onOpenAgenda} type="button"><LineIcon name="maintenance" /> Bekijk agenda</button>
        <button className="submitButton" onClick={onCreate} type="button"><LineIcon name="plus" /> Onderhoudstaak aanmaken</button>
      </section>
      <MaintenanceTaskList
        leveranciers={leveranciers}
        onderhoud={sortedOnderhoud}
        onComplete={onComplete}
        onOpenTask={onOpenTask}
        personen={personen}
      />
    </>
  );
}

function MaintenanceStats({ onderhoud }: { onderhoud: Onderhoud[] }) {
  const gepland = onderhoud.filter((taak) => taak.status === "Gepland" && !isOnderhoudAchterstallig(taak)).length;
  const inProces = onderhoud.filter((taak) => taak.status === "In proces" && !isOnderhoudAchterstallig(taak)).length;
  const voltooid = onderhoud.filter((taak) => taak.status === "Voltooid").length;
  const achterstallig = onderhoud.filter(isOnderhoudAchterstallig).length;

  return (
    <section className="maintenanceStats">
      <article><span>Totaal</span><strong>{onderhoud.length}</strong></article>
      <article><span>Gepland</span><strong>{gepland}</strong></article>
      <article><span>In proces</span><strong>{inProces}</strong></article>
      <article><span>Voltooid</span><strong>{voltooid}</strong></article>
      <article className={achterstallig ? "overdue" : ""}><span>Achterstallig</span><strong>{achterstallig}</strong></article>
    </section>
  );
}

function MaintenanceTaskList({
  onderhoud,
  personen,
  leveranciers,
  onComplete,
  onOpenTask,
}: {
  onderhoud: Onderhoud[];
  personen: Persoon[];
  leveranciers: Leverancier[];
  onComplete: (taakId: string) => void;
  onOpenTask: (taakId: string) => void;
}) {
  if (onderhoud.length === 0) {
    return <p className="emptyState">Geen onderhoud gepland.</p>;
  }

  return (
    <section className="maintenanceList">
      {onderhoud.map((taak) => (
        <article className={`maintenanceCard ${taak.status === "Voltooid" ? "completed" : ""}`} key={taak.id}>
          <button
            aria-label={taak.status === "Voltooid" ? "Taak terugzetten naar gepland" : "Taak als voltooid markeren"}
            className={`completeButton ${taak.status === "Voltooid" ? "done" : ""}`}
            onClick={(event) => {
              event.stopPropagation();
              onComplete(taak.id);
            }}
            type="button"
          >
            {taak.status === "Voltooid" ? String.fromCharCode(10003) : ""}
          </button>
          <button className="maintenanceSummary" onClick={() => onOpenTask(taak.id)} type="button">
            <span>
              <strong>{taak.title}</strong>
              <small>{taak.datumGepland || "Geen datum"} · {taak.herhaling === "Geen" ? "eenmalig" : taak.herhaling}</small>
              <small>Verantwoordelijke: {getResponsibleLabel(taak, personen, leveranciers)}</small>
              {taak.status === "Voltooid" && taak.datumUitgevoerd && <small className="doneText">Taak voltooid op {taak.datumUitgevoerd}</small>}
            </span>
            <MaintenanceStatusBadge taak={taak} />
          </button>
        </article>
      ))}
    </section>
  );
}

function OnderhoudAgendaView({
  onderhoud,
  personen,
  leveranciers,
  filterValue,
  statusFilter,
  onComplete,
  onCreate,
  onFilterChange,
  onOpenTask,
  onStatusFilterChange,
}: {
  onderhoud: Onderhoud[];
  personen: Persoon[];
  leveranciers: Leverancier[];
  filterValue: string;
  statusFilter: string;
  onComplete: (taakId: string) => void;
  onCreate: () => void;
  onFilterChange: (value: string) => void;
  onOpenTask: (taakId: string) => void;
  onStatusFilterChange: (value: string) => void;
}) {
  const filteredOnderhoud = onderhoud
    .filter((taak) => filterValue === "all" || getResponsibleValue(taak) === filterValue)
    .filter((taak) => statusFilter === "all" || getOnderhoudDisplayStatus(taak) === statusFilter)
    .sort((a, b) => (a.datumGepland || "9999").localeCompare(b.datumGepland || "9999"));

  return (
    <>
      <section className="filterBlock agendaFilters">
        <label>
          Filter op verantwoordelijke
          <select value={filterValue} onChange={(event) => onFilterChange(event.target.value)}>
            <option value="all">Alle verantwoordelijken</option>
            <optgroup label="Personen">
              {personen.map((persoon) => <option key={persoon.id} value={`Persoon:${persoon.id}`}>{persoon.title}</option>)}
            </optgroup>
            <optgroup label="Leveranciers">
              {leveranciers.map((leverancier) => <option key={leverancier.id} value={`Leverancier:${leverancier.id}`}>{leverancier.title}</option>)}
            </optgroup>
          </select>
        </label>
        <label>
          Filter op status
          <select value={statusFilter} onChange={(event) => onStatusFilterChange(event.target.value)}>
            <option value="all">Alle statussen</option>
            <option>Gepland</option>
            <option>In proces</option>
            <option>Voltooid</option>
            <option>Achterstallig</option>
          </select>
        </label>
      </section>
      <section className="maintenanceActions">
        <button className="submitButton" onClick={onCreate} type="button"><LineIcon name="plus" /> Onderhoudstaak aanmaken</button>
      </section>
      <MaintenanceTaskList
        leveranciers={leveranciers}
        onderhoud={filteredOnderhoud}
        onComplete={onComplete}
        onOpenTask={onOpenTask}
        personen={personen}
      />
    </>
  );
}

function OnderhoudDetailView({
  defaultResponsibleId,
  leveranciers,
  onCancel,
  onCompleteAndReturn,
  onCreate,
  onDelete,
  onEdit,
  onReschedule,
  personen,
  taak,
}: {
  defaultResponsibleId: string;
  leveranciers: Leverancier[];
  onCancel: () => void;
  onCompleteAndReturn: (taakId: string) => void;
  onDelete: (taakId: string) => void;
  onCreate: (formData: FormData) => void;
  onEdit: (taakId: string, formData: FormData) => void;
  onReschedule: (taak: Onderhoud) => void;
  personen: Persoon[];
  taak?: Onderhoud;
}) {
  return (
    <>
      {taak && (
        <section className="maintenanceActions detailActions">
          <button className="submitButton" onClick={() => onCompleteAndReturn(taak.id)} type="button">Taak voltooid</button>
          <button className="ghostButton" onClick={() => onReschedule(taak)} type="button">Opnieuw plannen</button>
          <button className="ghostButton dangerButton" onClick={() => onDelete(taak.id)} type="button">Verwijderen</button>
        </section>
      )}
      <OnderhoudForm
        defaultResponsibleId={defaultResponsibleId}
        key={taak?.id ?? "nieuw"}
        leveranciers={leveranciers}
        mode={taak ? "edit" : "create"}
        onCancel={onCancel}
        onSubmit={(formData) => (taak ? onEdit(taak.id, formData) : onCreate(formData))}
        personen={personen}
        taak={taak}
      />
    </>
  );
}

function OnderhoudForm({
  defaultResponsibleId,
  leveranciers,
  mode,
  onCancel,
  onSubmit,
  personen,
  taak,
}: {
  defaultResponsibleId?: string;
  leveranciers: Leverancier[];
  mode: "create" | "edit";
  onCancel?: () => void;
  onSubmit: (formData: FormData) => void;
  personen: Persoon[];
  taak?: Onderhoud;
}) {
  const [herhaling, setHerhaling] = useState<Onderhoud["herhaling"]>(taak?.herhaling ?? "Geen");

  function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const form = event.currentTarget;
    onSubmit(new FormData(form));
    if (mode === "create") form.reset();
  }

  return (
    <form className="entryForm" onSubmit={handleSubmit}>
      <h2><LineIcon name="maintenance" /> {mode === "edit" ? "Onderhoudstaak bewerken" : "Onderhoudstaak maken"}</h2>
      <label>Titel<input name="title" defaultValue={taak?.title ?? ""} placeholder="Bijvoorbeeld: messen controleren" required /></label>
      <label>Datum gepland<input name="datumGepland" type="date" defaultValue={taak?.datumGepland ?? ""} /></label>
      <label>Datum uitgevoerd<input name="datumUitgevoerd" type="date" defaultValue={taak?.datumUitgevoerd ?? ""} /></label>
      <label>Type onderhoud<select name="typeOnderhoud" defaultValue={taak?.typeOnderhoud ?? "Preventief"}><option>Preventief</option><option>Correctief</option><option>Keuring</option><option>Schoonmaak</option></select></label>
      <label>Verantwoordelijke
        <select name="verantwoordelijke" defaultValue={taak ? getResponsibleValue(taak) : `Persoon:${defaultResponsibleId || personen[0]?.id || ""}`}>
          <optgroup label="Personen">
            {personen.map((persoon) => <option key={persoon.id} value={`Persoon:${persoon.id}`}>{persoon.title}</option>)}
          </optgroup>
          <optgroup label="Leveranciers">
            {leveranciers.map((leverancier) => <option key={leverancier.id} value={`Leverancier:${leverancier.id}`}>{leverancier.title}</option>)}
          </optgroup>
        </select>
      </label>
      <label>Status<select name="status" defaultValue={taak?.status ?? "Gepland"}><option>Gepland</option><option>In proces</option><option>Voltooid</option></select></label>
      <label>Herhaling<select name="herhaling" value={herhaling} onChange={(event) => setHerhaling(event.target.value as Onderhoud["herhaling"])}><option>Geen</option><option>Wekelijks</option><option>Maandelijks</option><option>Jaarlijks</option><option>Eerste weekdag van de maand</option><option>Eerste weekdag van het kwartaal</option></select></label>
      {herhaling !== "Geen" && (
        <>
          <label>Eerste weekdag<select name="herhalingWeekdag" defaultValue={taak?.herhalingWeekdag ?? "Maandag"}>{weekdagen.map((weekdag) => <option key={weekdag}>{weekdag}</option>)}</select></label>
        </>
      )}
      <label>Opmerking<textarea name="opmerking" rows={3} defaultValue={taak?.opmerking ?? ""} placeholder="Wat moet er gebeuren?" /></label>
      <div className="formActions">
        {onCancel && <button className="ghostButton" onClick={onCancel} type="button">Terug</button>}
        <button className="submitButton gold" type="submit">{mode === "edit" ? "Wijzigingen opslaan" : "Onderhoudstaak opslaan"}</button>
      </div>
    </form>
  );
}

function MaintenanceStatusBadge({ taak }: { taak: Onderhoud }) {
  const status = getOnderhoudDisplayStatus(taak);
  return (
    <span className={`maintenanceStatus ${status === "Voltooid" ? "done" : status === "In proces" ? "busy" : status === "Achterstallig" ? "overdue" : ""}`}>
      {status}
    </span>
  );
}

function MachineAiTool({
  documents,
  machine,
  onderhoud,
  storingen,
}: {
  documents: MachineDocument[];
  machine: Machine;
  onderhoud: Onderhoud[];
  storingen: StoringOpmerking[];
}) {
  const [question, setQuestion] = useState("");
  const [answer, setAnswer] = useState("");
  const [sources, setSources] = useState<Array<{ documentType: string; title: string; url?: string }>>([]);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setBusy(true);
    setError("");
    setAnswer("");
    setSources([]);

    try {
      const response = await fetch("/api/ai-machine-question", {
        body: JSON.stringify({
          documents,
          machine,
          onderhoud,
          question,
          storingen,
        }),
        headers: {
          "Content-Type": "application/json",
        },
        method: "POST",
      });
      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.error || "AI-antwoord ophalen is niet gelukt.");
      }

      setAnswer(data.answer || "Geen antwoord ontvangen.");
      setSources(data.sources ?? []);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "AI-antwoord ophalen is niet gelukt.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <section className="aiTool">
      <div>
        <p className="eyebrow">AI-hulp</p>
        <h2><LineIcon name="spark" /> Vraag over deze machine</h2>
      </div>
      <form onSubmit={handleSubmit}>
        <label>
          Vraag of storing
          <textarea
            name="aiVraag"
            onChange={(event) => setQuestion(event.target.value)}
            placeholder="Bijvoorbeeld: waarom sealt de machine niet goed?"
            rows={3}
            value={question}
          />
        </label>
        <button className="submitButton red" disabled={busy || !question.trim()} type="submit">
          {busy ? "AI denkt mee..." : "Eerste oorzaken zoeken"}
        </button>
      </form>
      {error && <p className="aiAnswer error">{error}</p>}
      {answer && <AiAnswer answer={answer} />}
      {sources.length > 0 && (
        <div className="aiSources">
          <strong>Bronnen gebruikt</strong>
          {sources.map((source) => (
            source.url ? (
              <a href={source.url} key={`${source.documentType}-${source.title}`} rel="noreferrer" target="_blank">
                {source.documentType}: {source.title}
              </a>
            ) : (
              <span key={`${source.documentType}-${source.title}`}>
                {source.documentType}: {source.title}
              </span>
            )
          ))}
        </div>
      )}
    </section>
  );
}

function AiAnswer({ answer }: { answer: string }) {
  const headingPattern = /(##\s*(Kort antwoord|Eerste controles|Mogelijke oorzaken|Advies|Bronnen gebruikt))/gi;
  const normalized = answer.replace(headingPattern, "\n\n$1\n");
  const blocks = normalized.split(/\n{2,}/).map((block) => block.trim()).filter(Boolean);

  return (
    <div className="aiAnswer rich">
      {blocks.flatMap((block, blockIndex) => {
        const lines = block.split("\n").map((line) => line.trim()).filter(Boolean);

        if (lines[0]?.startsWith("##")) {
          const title = lines[0].replace(/^##\s+/, "");
          const content = lines.slice(1).join(" ");
          return [
            <h3 key={`${blockIndex}-heading`}>{title}</h3>,
            ...renderAiAnswerText(content, `${blockIndex}-content`),
          ];
        }

        return renderAiAnswerText(lines.join(" "), String(blockIndex));
      })}
    </div>
  );
}

function renderAiAnswerText(text: string, keyPrefix: string): ReactNode[] {
  const cleanText = text.trim();
  if (!cleanText) return [];
  const lines = cleanText.split("\n").map((line) => line.trim()).filter(Boolean);

  if (lines.length > 1) {
    return lines.flatMap((line, lineIndex) => renderAiAnswerText(line, `${keyPrefix}-${lineIndex}`));
  }

  if (cleanText.includes(" - ")) {
    const [intro, ...items] = cleanText.split(/\s+-\s+/).map((item) => item.trim()).filter(Boolean);
    return [
      intro && <p key={`${keyPrefix}-intro`}>{intro}</p>,
      items.length > 0 && (
        <ul key={`${keyPrefix}-list`}>
          {items.map((item, itemIndex) => <li key={itemIndex}>{item}</li>)}
        </ul>
      ),
    ].filter(Boolean);
  }

  if (lines.every((line) => line.startsWith("- "))) {
    return [
      <ul key={keyPrefix}>
        {lines.map((line, lineIndex) => <li key={lineIndex}>{line.replace(/^-\s+/, "")}</li>)}
      </ul>,
    ];
  }

  if (lines.every((line) => /^\d+[.)]\s+/.test(line))) {
    return [
      <ol key={keyPrefix}>
        {lines.map((line, lineIndex) => <li key={lineIndex}>{line.replace(/^\d+[.)]\s+/, "")}</li>)}
      </ol>,
    ];
  }

  return [<p key={keyPrefix}>{cleanText}</p>];
}

function StoringForm({
  mode,
  onSubmit,
  storing,
}: {
  mode: "create" | "edit";
  onSubmit: (formData: FormData) => Promise<void>;
  storing?: StoringOpmerking;
}) {
  const [busy, setBusy] = useState(false);

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const form = event.currentTarget;
    setBusy(true);
    await onSubmit(new FormData(form));
    if (mode === "create") form.reset();
    setBusy(false);
  }

  return (
    <form className="entryForm alertForm" onSubmit={handleSubmit}>
      <h2><LineIcon name="alert" /> {mode === "edit" ? "Storing bekijken / bewerken" : "Nieuwe melding"}</h2>
      {storing?.datum && <p className="formMeta">Gemeld op {storing.datum}</p>}
      <label>Titel<input name="title" defaultValue={storing?.title ?? ""} placeholder="Bijvoorbeeld: afwijkend geluid" required /></label>
      <label>Type<select name="type" defaultValue={storing?.type ?? "Storing"}><option>Storing</option><option>Opmerking</option><option>Verbeterpunt</option></select></label>
      <label>Prioriteit<select name="prioriteit" defaultValue={storing?.prioriteit ?? "Normaal"}><option>Laag</option><option>Normaal</option><option>Hoog</option></select></label>
      <label>Status<select name="status" defaultValue={storing?.status ?? "Open"}><option>Open</option><option>In behandeling</option><option>Opgelost</option></select></label>
      <label>Omschrijving<textarea name="omschrijving" rows={3} defaultValue={storing?.omschrijving ?? ""} placeholder="Beschrijf wat je ziet of hoort" required /></label>
      <label>Oplossing<textarea name="oplossing" rows={3} defaultValue={storing?.oplossing ?? ""} placeholder="Wat is er gedaan of wat is de oplossing?" /></label>
      {storing?.bijlagen && storing.bijlagen.length > 0 && (
        <div className="attachmentGrid">
          {storing.bijlagen.map((bijlage) => (
            <a className="attachment" href={bijlage.dataUrl} key={bijlage.id} download={bijlage.naam}>
              {bijlage.type.startsWith("image/") ? (
                <span
                  aria-label={bijlage.naam}
                  className="attachmentPreview"
                  role="img"
                  style={{ backgroundImage: `url(${bijlage.dataUrl})` }}
                />
              ) : (
                <LineIcon name="document" />
              )}
              <span>{bijlage.naam}</span>
            </a>
          ))}
        </div>
      )}
      <label className="fileBox red"><LineIcon name="camera" /> Foto of bijlage toevoegen<input accept="image/*,.pdf,.doc,.docx" capture="environment" multiple name="bijlagen" type="file" /></label>
      <button className="submitButton red" disabled={busy} type="submit">{busy ? "Opslaan..." : mode === "edit" ? "Wijzigingen opslaan" : "Melding opslaan"}</button>
    </form>
  );
}

function DocumentDetail({
  document,
  onBack,
  onSubmit,
}: {
  document: MachineDocument;
  onBack: () => void;
  onSubmit: (formData: FormData) => void;
}) {
  function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const form = event.currentTarget;
    onSubmit(new FormData(form));
  }

  const documentUrl = document.bestandDataUrl || document.url;
  const canPreview = documentUrl && documentUrl !== "#";

  return (
    <div className="documentDetail">
      <button className="smallButton" onClick={onBack} type="button"><LineIcon name="back" /> Terug naar documenten</button>
      <section className="documentPreview">
        <h2><LineIcon name="document" /> {document.title}</h2>
        {canPreview ? (
          <div className="documentOpenCard">
            <LineIcon name="document" />
            <div>
              <strong>{document.title}</strong>
              <p>{document.documentType} - {document.omschrijving || "Geen omschrijving"}</p>
              {document.vervaldatum && <small>Vervaldatum: {document.vervaldatum}</small>}
            </div>
            <a className="submitButton gold" href={documentUrl} target="_blank" rel="noreferrer">
              Openen in SharePoint
            </a>
          </div>
        ) : (
          <p className="emptyState">Dit testdocument heeft nog geen gekoppeld bestand.</p>
        )}
      </section>
      <form className="entryForm" onSubmit={handleSubmit}>
        <h2><LineIcon name="document" /> Documentvelden bewerken</h2>
        <label>Titel<input name="title" defaultValue={document.title} required /></label>
        <label>Document type<select name="documentType" defaultValue={document.documentType}><option>Handleiding</option><option>Keuring</option><option>Onderhoud</option><option>Foto</option><option>Overig</option></select></label>
        <label>Omschrijving<textarea name="omschrijving" rows={3} defaultValue={document.omschrijving} /></label>
        <label>Vervaldatum<input name="vervaldatum" type="date" defaultValue={document.vervaldatum} /></label>
        <label className="checkLine"><input name="actief" type="checkbox" defaultChecked={document.actief} /> Actief</label>
        <button className="submitButton gold" type="submit">Document opslaan</button>
      </form>
    </div>
  );
}

function DocumentList({ documenten, onOpen }: { documenten: MachineDocument[]; onOpen: (documentId: string) => void }) {
  if (documenten.length === 0) return <p className="emptyState">Geen documenten.</p>;

  return (
    <div className="recordList">
      {documenten.map((doc) => (
        <article className="recordCard" key={doc.id}>
          <strong>{doc.title}</strong>
          <p>{doc.documentType} - {doc.omschrijving || "Geen omschrijving"}</p>
          <button className="smallButton" onClick={() => onOpen(doc.id)} type="button">
            <LineIcon name="document" />
            Openen / bewerken
          </button>
          {doc.bestandDataUrl && <a className="fileLink" href={doc.bestandDataUrl} download={doc.bestandNaam}>Download {doc.bestandNaam}</a>}
        </article>
      ))}
    </div>
  );
}

function StoringList({
  onOpen,
  onResolve,
  storingen,
}: {
  onOpen?: (storingId: string) => void;
  onResolve?: (storingId: string) => void;
  storingen: StoringOpmerking[];
}) {
  if (storingen.length === 0) return <p className="emptyState">Geen meldingen.</p>;

  return (
    <div className="recordList">
      {storingen.map((item) => (
        <article className={`recordCard alertCard ${item.status === "Opgelost" ? "resolved" : ""}`} key={item.id}>
          <div className="alertCardHeader">
            <span className={`alertStatus ${item.status === "Opgelost" ? "resolved" : ""}`}>
              {item.status === "Opgelost" ? `${String.fromCharCode(10003)} Opgelost` : item.status}
            </span>
            <div className="alertCardActions">
              {onOpen && (
                <button className="smallButton" onClick={() => onOpen(item.id)} type="button">
                  Openen
                </button>
              )}
              {onResolve && item.status !== "Opgelost" && (
                <button className="smallButton resolveButton" onClick={() => onResolve(item.id)} type="button">
                  {String.fromCharCode(10003)} Opgelost
                </button>
              )}
            </div>
          </div>
          <strong>{item.title}</strong>
          <p>{item.type} - {item.prioriteit} - {item.omschrijving}</p>
          {item.oplossing && <p className="resolvedText">{item.oplossing}</p>}
          {item.bijlagen && item.bijlagen.length > 0 && (
            <div className="attachmentGrid">
              {item.bijlagen.map((bijlage) => (
                <a className="attachment" href={bijlage.dataUrl} key={bijlage.id} download={bijlage.naam}>
                  {bijlage.type.startsWith("image/") ? (
                    <span
                      aria-label={bijlage.naam}
                      className="attachmentPreview"
                      role="img"
                      style={{ backgroundImage: `url(${bijlage.dataUrl})` }}
                    />
                  ) : (
                    <LineIcon name="document" />
                  )}
                  <span>{bijlage.naam}</span>
                </a>
              ))}
            </div>
          )}
        </article>
      ))}
    </div>
  );
}

function LineIcon({ name }: { name: IconName }) {
  const common = { fill: "none", stroke: "currentColor", strokeLinecap: "round" as const, strokeLinejoin: "round" as const, strokeWidth: 2 };
  return (
    <svg aria-hidden="true" className="lineIcon" viewBox="0 0 24 24">
      {name === "home" && <><path {...common} d="M4 11.5 12 5l8 6.5" /><path {...common} d="M6.5 10.5V20h11v-9.5" /><path {...common} d="M10 20v-5h4v5" /></>}
      {name === "department" && <><path {...common} d="M5 20V8l7-3 7 3v12" /><path {...common} d="M8 20v-5h8v5" /><path {...common} d="M9 10h.01M12 10h.01M15 10h.01" /></>}
      {name === "person" && <><circle {...common} cx="12" cy="8" r="3" /><path {...common} d="M5 20c1.5-4 4-6 7-6s5.5 2 7 6" /></>}
      {name === "machine" && <><rect {...common} x="4" y="8" width="16" height="10" rx="2" /><path {...common} d="M8 8V5h8v3M8 18v2M16 18v2" /><path {...common} d="M8 13h3M14 13h2" /></>}
      {name === "document" && <><path {...common} d="M7 3h7l4 4v14H7z" /><path {...common} d="M14 3v5h4M9 12h6M9 16h6" /></>}
      {name === "maintenance" && <><path {...common} d="M14.5 6.5 17 4l3 3-2.5 2.5" /><path {...common} d="m4 20 8.5-8.5" /><path {...common} d="M6 6h5M6 10h3M4 4l16 16" /></>}
      {name === "alert" && <><path {...common} d="M12 4 3.5 19h17z" /><path {...common} d="M12 9v4M12 16h.01" /></>}
      {name === "camera" && <><path {...common} d="M4 8h4l1.5-2h5L16 8h4v11H4z" /><circle {...common} cx="12" cy="13.5" r="3" /></>}
      {name === "plus" && <><path {...common} d="M12 5v14M5 12h14" /></>}
      {name === "back" && <><path {...common} d="M15 6 9 12l6 6" /><path {...common} d="M10 12h10" /></>}
      {name === "spark" && <><path {...common} d="M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8z" /><path {...common} d="M18 15l.8 2.2L21 18l-2.2.8L18 21l-.8-2.2L15 18l2.2-.8z" /></>}
    </svg>
  );
}

function BeheerView({
  data,
  activeEntity,
  setActiveEntity,
  updateEntity,
}: {
  data: AppData;
  activeEntity: EntityName;
  setActiveEntity: (entity: EntityName) => void;
  updateEntity: <T extends EntityName>(entity: T, records: AppData[T]) => void;
}) {
  const records = data[activeEntity] as Array<Record<string, unknown>>;

  function addRecord() {
    const template = records[0] ?? { id: "", title: "" };
    const nextRecord = Object.fromEntries(
      Object.entries(template).map(([key, value]) => {
        if (key === "id") return [key, createId(activeEntity.slice(0, 3))];
        if (Array.isArray(value)) return [key, []];
        if (typeof value === "boolean") return [key, true];
        if (typeof value === "number") return [key, records.length + 1];
        return [key, ""];
      }),
    );

    updateEntity(activeEntity, [...records, nextRecord] as never);
  }

  function updateRecord(index: number, key: string, value: string | boolean | number) {
    const nextRecords = records.map((record, recordIndex) =>
      recordIndex === index ? { ...record, [key]: value } : record,
    );

    updateEntity(activeEntity, nextRecords as never);
  }

  return (
    <div className="beheerLayout">
      <aside className="panel beheerMenu">
        <div className="sectionHeader">
          <div>
            <p className="eyebrow">Hub</p>
            <h2>Beheer</h2>
          </div>
        </div>
        <PilotNotice />
        <div className="buttonList compact">
          {entityOrder.map((entity) => (
            <button className={entity === activeEntity ? "choice active" : "choice"} key={entity} onClick={() => setActiveEntity(entity)} type="button">
              <strong>{entityLabels[entity]}</strong>
              <span>{data[entity].length} records</span>
            </button>
          ))}
        </div>
      </aside>

      <section className="panel tablePanel">
        <div className="sectionHeader">
          <div>
            <p className="eyebrow">Lokale testdata</p>
            <h2>{entityLabels[activeEntity]}</h2>
          </div>
          <button className="primaryButton" onClick={addRecord} type="button"><LineIcon name="plus" /> Record toevoegen</button>
        </div>
        <div className="tableWrap">
          <table>
            <thead><tr>{Object.keys(records[0] ?? { title: "" }).map((key) => <th key={key}>{key}</th>)}</tr></thead>
            <tbody>
              {records.map((record, index) => (
                <tr key={String(record.id ?? index)}>
                  {Object.entries(record).map(([key, value]) => (
                    <td key={key}><EditableCell fieldKey={key} onChange={(nextValue) => updateRecord(index, key, nextValue)} value={value} /></td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}

function EditableCell({ fieldKey, value, onChange }: { fieldKey: string; value: unknown; onChange: (value: string | boolean | number) => void }) {
  if (fieldKey === "id" || Array.isArray(value) || (typeof value === "string" && value.startsWith("data:"))) return <span className="idText">{Array.isArray(value) ? `${value.length} bijlage(n)` : String(value)}</span>;
  if (typeof value === "boolean") return <input aria-label={fieldKey} checked={value} onChange={(event) => onChange(event.target.checked)} type="checkbox" />;
  if (typeof value === "number") return <input aria-label={fieldKey} onChange={(event) => onChange(Number(event.target.value))} type="number" value={value} />;
  return <input aria-label={fieldKey} onChange={(event) => onChange(event.target.value)} value={String(value ?? "")} />;
}


