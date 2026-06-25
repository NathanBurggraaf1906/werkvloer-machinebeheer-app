import { NextResponse } from "next/server";
import type { Machine, MachineDocument, Onderhoud, StoringOpmerking } from "@/lib/types";

type RequestBody = {
  documents: MachineDocument[];
  machine: Machine;
  onderhoud: Onderhoud[];
  question: string;
  storingen: StoringOpmerking[];
};

function selectDocuments(question: string, documents: MachineDocument[]) {
  const normalizedQuestion = question.toLowerCase();
  const activeDocuments = documents.filter((document) => document.actief);
  const selected = activeDocuments.filter((document) => {
    if (document.documentType === "Handleiding") return true;
    if (normalizedQuestion.match(/onderhoud|schoonmaak|keuring|service|periodiek/)) {
      return document.documentType === "Onderhoud" || document.documentType === "Keuring";
    }
    if (normalizedQuestion.match(/garantie|leverancier|certificaat|keuren/)) {
      return document.documentType === "Keuring" || document.documentType === "Overig";
    }
    return false;
  });

  return (selected.length ? selected : activeDocuments)
    .filter((document) => document.documentType !== "Foto")
    .slice(0, 8);
}

function compactDocument(document: MachineDocument) {
  return {
    documentType: document.documentType,
    omschrijving: document.omschrijving,
    title: document.title,
    url: document.url?.startsWith("data:") ? "" : document.url,
    vervaldatum: document.vervaldatum,
  };
}

export async function POST(request: Request) {
  const apiKey = process.env.OPENAI_API_KEY;

  if (!apiKey) {
    return NextResponse.json(
      { error: "OPENAI_API_KEY ontbreekt nog op de server/Vercel environment." },
      { status: 500 },
    );
  }

  const body = (await request.json()) as RequestBody;
  const question = body.question?.trim();

  if (!question) {
    return NextResponse.json({ error: "Vraag ontbreekt." }, { status: 400 });
  }

  const documents = selectDocuments(question, body.documents ?? []);
  const openStoringen = (body.storingen ?? []).filter((storing) => storing.status !== "Opgelost").slice(0, 8);
  const recentOnderhoud = [...(body.onderhoud ?? [])]
    .sort((a, b) => (b.datumGepland || "").localeCompare(a.datumGepland || ""))
    .slice(0, 8);

  const promptData = {
    documentenEerstGebruiken: documents.map(compactDocument),
    machine: body.machine,
    openStoringen,
    recentOnderhoud,
    vraag: question,
  };

  const response = await fetch("https://api.openai.com/v1/responses", {
    body: JSON.stringify({
      input: [
        {
          content:
            "Je bent een praktische werkvloer-assistent voor machinebeheer bij 1906 Makers van Charcuterie. " +
            "Gebruik eerst en vooral de gekoppelde documenten als bron. Als documentinhoud ontbreekt en alleen titel/omschrijving/link beschikbaar is, zeg dat expliciet. " +
            "Geef korte, veilige en concrete stappen. Adviseer bij gevaar, twijfel of ingrijpende reparatie altijd om de verantwoordelijke of leverancier te betrekken. " +
            "Antwoord in het Nederlands in nette Markdown-opmaak. Gebruik exact deze koppen: ## Kort antwoord, ## Eerste controles, ## Mogelijke oorzaken, ## Advies, ## Bronnen gebruikt. " +
            "Gebruik korte alinea's en bullets met '- '. Geen lange tekstblokken.",
          role: "system",
        },
        {
          content: JSON.stringify(promptData),
          role: "user",
        },
      ],
      max_output_tokens: 700,
      model: process.env.OPENAI_MODEL || "gpt-4.1-mini",
    }),
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    method: "POST",
  });

  if (!response.ok) {
    const errorText = await response.text();
    return NextResponse.json(
      { error: `OpenAI fout ${response.status}: ${errorText || response.statusText}` },
      { status: 500 },
    );
  }

  const data = await response.json();
  const answer =
    data.output_text ||
    data.output?.flatMap((item: { content?: Array<{ text?: string }> }) => item.content ?? [])
      .map((content: { text?: string }) => content.text)
      .filter(Boolean)
      .join("\n") ||
    "Geen antwoord ontvangen.";

  return NextResponse.json({
    answer,
    sources: documents.map((document) => ({
      documentType: document.documentType,
      title: document.title,
      url: document.url?.startsWith("data:") ? "" : document.url,
    })),
  });
}
