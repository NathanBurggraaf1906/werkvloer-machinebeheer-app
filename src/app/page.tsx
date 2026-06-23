"use client";

import Image from "next/image";
import { FormEvent, useEffect, useMemo, useState } from "react";
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

const storageKey = "werkvloer-machinebeheer-v1";

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
  | "storingen";
type IconName = "home" | "department" | "person" | "machine" | "document" | "maintenance" | "alert" | "camera" | "plus" | "back" | "spark";

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
  const { data, updateEntity, resetData } = useAppData();
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
          <div className="navSwitch">
            <button
              className={section === "werkvloer" ? "active" : ""}
              onClick={goHome}
              type="button"
            >
              Werkvloer
            </button>
            <button
              className={section === "beheer" ? "active" : ""}
              onClick={() => setSection("beheer")}
              type="button"
            >
              Beheer
            </button>
          </div>
        </div>
      </header>

      {section === "werkvloer" ? (
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
          updateDocumenten={(documenten) => updateEntity("documenten", documenten)}
          updateOnderhoud={(onderhoud) => updateEntity("onderhoud", onderhoud)}
          updateStoringen={(storingen) => updateEntity("storingen", storingen)}
        />
      ) : (
        <BeheerView
          activeEntity={activeEntity}
          data={data}
          resetData={resetData}
          setActiveEntity={setActiveEntity}
          updateEntity={updateEntity}
        />
      )}
    </main>
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
  updateDocumenten,
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
  updateDocumenten: (documenten: MachineDocument[]) => void;
  updateOnderhoud: (onderhoud: Onderhoud[]) => void;
  updateStoringen: (storingen: StoringOpmerking[]) => void;
}) {
  const [selectedDocumentId, setSelectedDocumentId] = useState("");
  const [editingOnderhoudId, setEditingOnderhoudId] = useState("");
  const [maintenanceFilter, setMaintenanceFilter] = useState("all");
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
    const responsible = parseResponsibleValue(String(formData.get("verantwoordelijke") || `Persoon:${data.personen[0]?.id ?? ""}`));

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
      herhalingTot: String(formData.get("herhalingTot") || ""),
      opmerking: String(formData.get("opmerking") || ""),
    };

    updateOnderhoud([taak, ...data.onderhoud]);
    setEditingOnderhoudId(taak.id);
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
            herhalingTot: String(formData.get("herhalingTot") || ""),
            opmerking: String(formData.get("opmerking") || ""),
          }
        : taak,
    );

    updateOnderhoud(bijgewerkt);
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
  }

  if (flowScreen === "start") {
    return (
      <section className="mobileScreen heroScreen">
        <div className="heroCopy">
          <p className="eyebrow">Mobiele werkvloer-app</p>
          <h1>Machinebeheer zonder omwegen.</h1>
          <p>
            Kies een afdeling, open een machinepaspoort en registreer snel wat er
            speelt op de werkvloer.
          </p>
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
        <div className="passportCard">
          <span className={`statusPill ${selectedMachine.status === "Operationeel" ? "good" : "alert"}`}>{selectedMachine.status}</span>
          <p>{selectedMachine.omschrijving}</p>
          <dl className="passportMeta">
            <div><dt>Serie</dt><dd>{selectedMachine.serieNummer}</dd></div>
            <div><dt>Verantwoordelijke</dt><dd>{verantwoordelijke?.title ?? "-"}</dd></div>
            <div><dt>Leverancier</dt><dd>{leverancier?.title ?? "-"}</dd></div>
            <div><dt>Garantie</dt><dd>{selectedMachine.garantieVerloopdatum || "-"}</dd></div>
          </dl>
        </div>
        <div className="actionGrid">
          <ActionButton icon="document" label="Documenten" onClick={() => setFlowScreen("documenten")} />
          <ActionButton icon="maintenance" label="Onderhoud" onClick={() => setFlowScreen("onderhoud")} tone="gold" />
          <ActionButton icon="alert" label="Storingen" onClick={() => setFlowScreen("storingen")} tone="red" />
        </div>
        <MachineAiTool machine={selectedMachine} storingen={machineStoringen} />
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
            onderhoud={machineOnderhoud}
            personen={data.personen}
            leveranciers={data.leveranciers}
            editingTaak={editingOnderhoud?.machineId === selectedMachine.id ? editingOnderhoud : undefined}
            filterValue={maintenanceFilter}
            onCreate={addOnderhoud}
            onEdit={editOnderhoud}
            onFilterChange={setMaintenanceFilter}
            onSelectEdit={setEditingOnderhoudId}
            onStopEdit={() => setEditingOnderhoudId("")}
          />
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
      {selectedMachine && <StoringForm onSubmit={addStoring} />}
      <StoringList storingen={selectedMachine ? machineStoringen : data.storingen} />
    </section>
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

function DocumentUploadForm({ onSubmit }: { onSubmit: (formData: FormData) => Promise<void> }) {
  const [busy, setBusy] = useState(false);

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setBusy(true);
    await onSubmit(new FormData(event.currentTarget));
    event.currentTarget.reset();
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

function getOnderhoudAgendaItems(onderhoud: Onderhoud[]) {
  const items: Array<{ id: string; datum: string; title: string; status: Onderhoud["status"]; herhaling: Onderhoud["herhaling"] }> = [];

  for (const taak of onderhoud) {
    if (!taak.datumGepland) continue;

    const start = new Date(`${taak.datumGepland}T00:00:00`);
    const eind = taak.herhalingTot ? new Date(`${taak.herhalingTot}T00:00:00`) : undefined;
    const maxItems = taak.herhaling === "Geen" ? 1 : 6;

    for (let index = 0; index < maxItems; index += 1) {
      const datum = new Date(start);

      if (taak.herhaling === "Wekelijks") datum.setDate(start.getDate() + index * 7);
      if (taak.herhaling === "Maandelijks") datum.setMonth(start.getMonth() + index);
      if (taak.herhaling === "Jaarlijks") datum.setFullYear(start.getFullYear() + index);
      if (taak.herhaling === "Eerste weekdag van de maand") {
        datum.setTime(getFirstWeekdayDate(start.getFullYear(), start.getMonth() + index, taak.herhalingWeekdag ?? "Maandag").getTime());
      }
      if (taak.herhaling === "Eerste weekdag van het kwartaal") {
        const startKwartaalMaand = Math.floor(start.getMonth() / 3) * 3;
        datum.setTime(getFirstWeekdayDate(start.getFullYear(), startKwartaalMaand + index * 3, taak.herhalingWeekdag ?? "Maandag").getTime());
      }

      if (eind && datum > eind) break;

      items.push({
        id: `${taak.id}-${index}`,
        datum: datum.toISOString().slice(0, 10),
        herhaling: taak.herhaling,
        status: taak.status,
        title: taak.title,
      });
    }
  }

  return items.sort((a, b) => a.datum.localeCompare(b.datum));
}

function OnderhoudPanel({
  onderhoud,
  personen,
  leveranciers,
  editingTaak,
  filterValue,
  onCreate,
  onEdit,
  onFilterChange,
  onSelectEdit,
  onStopEdit,
}: {
  onderhoud: Onderhoud[];
  personen: Persoon[];
  leveranciers: Leverancier[];
  editingTaak?: Onderhoud;
  filterValue: string;
  onCreate: (formData: FormData) => void;
  onEdit: (taakId: string, formData: FormData) => void;
  onFilterChange: (value: string) => void;
  onSelectEdit: (taakId: string) => void;
  onStopEdit: () => void;
}) {
  const filteredOnderhoud = filterValue === "all"
    ? onderhoud
    : onderhoud.filter((taak) => getResponsibleValue(taak) === filterValue);
  const agendaItems = getOnderhoudAgendaItems(filteredOnderhoud);

  return (
    <>
      <OnderhoudForm
        key={editingTaak?.id ?? "nieuw"}
        leveranciers={leveranciers}
        mode={editingTaak ? "edit" : "create"}
        onCancel={editingTaak ? onStopEdit : undefined}
        onSubmit={(formData) => (editingTaak ? onEdit(editingTaak.id, formData) : onCreate(formData))}
        personen={personen}
        taak={editingTaak}
      />
      <section className="filterBlock">
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
      </section>
      <section className="agendaBlock">
        <h2><LineIcon name="maintenance" /> Agenda</h2>
        {agendaItems.length > 0 ? (
          <div className="agendaList">
            {agendaItems.map((item) => (
              <article className="agendaItem" key={item.id}>
                <StatusIcon status={item.status} />
                <time>{item.datum}</time>
                <strong>{item.title}</strong>
                <span className={`maintenanceStatus ${item.status === "Voltooid" ? "done" : item.status === "In proces" ? "busy" : ""}`}>{item.status}</span>
                <small>{item.herhaling === "Geen" ? "Eenmalig" : item.herhaling}</small>
              </article>
            ))}
          </div>
        ) : (
          <p className="emptyState">Geen onderhoud in de agenda.</p>
        )}
      </section>
      <div className="maintenanceList">
        {filteredOnderhoud.length > 0 ? (
          filteredOnderhoud.map((taak) => (
            <article className={editingTaak?.id === taak.id ? "maintenanceCard active" : "maintenanceCard"} key={taak.id}>
              <button className="maintenanceSummary" onClick={() => onSelectEdit(taak.id)} type="button">
                <StatusIcon status={taak.status} />
                <span>
                  <strong>{taak.title}</strong>
                  <small>{taak.datumGepland || "Geen datum"} - {taak.herhaling === "Geen" ? "eenmalig" : taak.herhaling}</small>
                  <small>Verantwoordelijke: {getResponsibleLabel(taak, personen, leveranciers)}</small>
                </span>
                <span className={`maintenanceStatus ${taak.status === "Voltooid" ? "done" : taak.status === "In proces" ? "busy" : ""}`}>{taak.status}</span>
              </button>
            </article>
          ))
        ) : (
          <p className="emptyState">Geen onderhoud gepland.</p>
        )}
      </div>
    </>
  );
}

function OnderhoudForm({
  leveranciers,
  mode,
  onCancel,
  onSubmit,
  personen,
  taak,
}: {
  leveranciers: Leverancier[];
  mode: "create" | "edit";
  onCancel?: () => void;
  onSubmit: (formData: FormData) => void;
  personen: Persoon[];
  taak?: Onderhoud;
}) {
  function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    onSubmit(new FormData(event.currentTarget));
    if (mode === "create") event.currentTarget.reset();
  }

  return (
    <form className="entryForm" onSubmit={handleSubmit}>
      <h2><LineIcon name="maintenance" /> {mode === "edit" ? "Onderhoudstaak bewerken" : "Onderhoudstaak maken"}</h2>
      <label>Titel<input name="title" defaultValue={taak?.title ?? ""} placeholder="Bijvoorbeeld: messen controleren" required /></label>
      <label>Datum gepland<input name="datumGepland" type="date" defaultValue={taak?.datumGepland ?? ""} /></label>
      <label>Datum uitgevoerd<input name="datumUitgevoerd" type="date" defaultValue={taak?.datumUitgevoerd ?? ""} /></label>
      <label>Type onderhoud<select name="typeOnderhoud" defaultValue={taak?.typeOnderhoud ?? "Preventief"}><option>Preventief</option><option>Correctief</option><option>Keuring</option><option>Schoonmaak</option></select></label>
      <label>Verantwoordelijke
        <select name="verantwoordelijke" defaultValue={taak ? getResponsibleValue(taak) : `Persoon:${personen[0]?.id ?? ""}`}>
          <optgroup label="Personen">
            {personen.map((persoon) => <option key={persoon.id} value={`Persoon:${persoon.id}`}>{persoon.title}</option>)}
          </optgroup>
          <optgroup label="Leveranciers">
            {leveranciers.map((leverancier) => <option key={leverancier.id} value={`Leverancier:${leverancier.id}`}>{leverancier.title}</option>)}
          </optgroup>
        </select>
      </label>
      <label>Status<select name="status" defaultValue={taak?.status ?? "Gepland"}><option>Gepland</option><option>In proces</option><option>Voltooid</option></select></label>
      <label>Herhaling<select name="herhaling" defaultValue={taak?.herhaling ?? "Geen"}><option>Geen</option><option>Wekelijks</option><option>Maandelijks</option><option>Jaarlijks</option><option>Eerste weekdag van de maand</option><option>Eerste weekdag van het kwartaal</option></select></label>
      <label>Eerste weekdag<select name="herhalingWeekdag" defaultValue={taak?.herhalingWeekdag ?? "Maandag"}>{weekdagen.map((weekdag) => <option key={weekdag}>{weekdag}</option>)}</select></label>
      <label>Herhalen tot<input name="herhalingTot" type="date" defaultValue={taak?.herhalingTot ?? ""} /></label>
      <label>Opmerking<textarea name="opmerking" rows={3} defaultValue={taak?.opmerking ?? ""} placeholder="Wat moet er gebeuren?" /></label>
      <div className="formActions">
        {onCancel && <button className="ghostButton" onClick={onCancel} type="button">Nieuwe taak</button>}
        <button className="submitButton gold" type="submit">{mode === "edit" ? "Wijzigingen opslaan" : "Onderhoudstaak opslaan"}</button>
      </div>
    </form>
  );
}

function StatusIcon({ status }: { status: Onderhoud["status"] }) {
  return (
    <span className={`statusIcon ${status === "Voltooid" ? "done" : status === "In proces" ? "busy" : "planned"}`} title={status}>
      <LineIcon name="maintenance" />
    </span>
  );
}

function MachineAiTool({ machine, storingen }: { machine: Machine; storingen: StoringOpmerking[] }) {
  const [question, setQuestion] = useState("");
  const [answer, setAnswer] = useState("");

  function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const openStoringen = storingen.filter((storing) => storing.status !== "Opgelost");
    const storingHint = openStoringen.length > 0
      ? `Er staan ${openStoringen.length} open melding(en). Controleer eerst de laatst gemelde symptomen, veiligheid en basisinstellingen.`
      : "Er staan geen open meldingen bij deze machine in de testdata.";

    setAnswer(
      `Mock AI-antwoord voor ${machine.title}: ${storingHint} Mogelijke eerste stappen: machine veiligstellen, visuele controle uitvoeren, handleiding/documenten raadplegen en verantwoordelijke inschakelen. Later koppelen we dit blok aan een echte AI-endpoint met machinegegevens, documenten en storingshistorie.`,
    );
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
        <button className="submitButton red" type="submit">Eerste oorzaken zoeken</button>
      </form>
      {answer && <p className="aiAnswer">{answer}</p>}
    </section>
  );
}

function StoringForm({ onSubmit }: { onSubmit: (formData: FormData) => Promise<void> }) {
  const [busy, setBusy] = useState(false);

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setBusy(true);
    await onSubmit(new FormData(event.currentTarget));
    event.currentTarget.reset();
    setBusy(false);
  }

  return (
    <form className="entryForm alertForm" onSubmit={handleSubmit}>
      <h2><LineIcon name="alert" /> Nieuwe melding</h2>
      <label>Titel<input name="title" placeholder="Bijvoorbeeld: afwijkend geluid" required /></label>
      <label>Type<select name="type" defaultValue="Storing"><option>Storing</option><option>Opmerking</option><option>Verbeterpunt</option></select></label>
      <label>Prioriteit<select name="prioriteit" defaultValue="Normaal"><option>Laag</option><option>Normaal</option><option>Hoog</option></select></label>
      <label>Omschrijving<textarea name="omschrijving" rows={3} placeholder="Beschrijf wat je ziet of hoort" required /></label>
      <label className="fileBox red"><LineIcon name="camera" /> Foto of bijlage toevoegen<input accept="image/*,.pdf,.doc,.docx" capture="environment" multiple name="bijlagen" type="file" /></label>
      <button className="submitButton red" disabled={busy} type="submit">{busy ? "Opslaan..." : "Melding opslaan"}</button>
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
    onSubmit(new FormData(event.currentTarget));
  }

  const documentUrl = document.bestandDataUrl || document.url;
  const canPreview = documentUrl && documentUrl !== "#";

  return (
    <div className="documentDetail">
      <button className="smallButton" onClick={onBack} type="button"><LineIcon name="back" /> Terug naar documenten</button>
      <section className="documentPreview">
        <h2><LineIcon name="document" /> {document.title}</h2>
        {canPreview ? (
          <>
            <iframe className="documentFrame" src={documentUrl} title={document.title} />
            <a className="fileLink" href={documentUrl} target="_blank" rel="noreferrer">
              Open in nieuw venster
            </a>
          </>
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

function StoringList({ storingen }: { storingen: StoringOpmerking[] }) {
  if (storingen.length === 0) return <p className="emptyState">Geen meldingen.</p>;

  return (
    <div className="recordList">
      {storingen.map((item) => (
        <article className="recordCard alertCard" key={item.id}>
          <strong>{item.title}</strong>
          <p>{item.prioriteit} - {item.status} - {item.omschrijving}</p>
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
  resetData,
}: {
  data: AppData;
  activeEntity: EntityName;
  setActiveEntity: (entity: EntityName) => void;
  updateEntity: <T extends EntityName>(entity: T, records: AppData[T]) => void;
  resetData: () => void;
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
        <div className="buttonList compact">
          {entityOrder.map((entity) => (
            <button className={entity === activeEntity ? "choice active" : "choice"} key={entity} onClick={() => setActiveEntity(entity)} type="button">
              <strong>{entityLabels[entity]}</strong>
              <span>{data[entity].length} records</span>
            </button>
          ))}
        </div>
        <button className="ghostButton" onClick={resetData} type="button">Testdata herstellen</button>
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


