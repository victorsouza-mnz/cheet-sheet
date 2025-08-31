const authBtn = document.getElementById("auth");
const sheetsList = document.getElementById("sheets");
const sheetPanel = document.getElementById("sheetPanel");
const sheetTitle = document.getElementById("sheetTitle");
const sheetLink = document.getElementById("sheetLink");
const sheetPrompt = document.getElementById("sheetPrompt");
const sheetResult = document.getElementById("sheetResult");
const backBtn = document.getElementById("backBtn");
const runSheetAI = document.getElementById("runSheetAI");
const groqKeyInput = document.getElementById("groqKey");
const saveGroqKeyBtn = document.getElementById("saveGroqKey");
const mainPanel = document.getElementById("mainPanel");

let currentSheets = [];
let selectedSheet = null;

// Event Listeners
authBtn.addEventListener("click", handleAuth);
sheetsList.addEventListener("click", handleSheetClick);
backBtn.addEventListener("click", handleBack);
runSheetAI.addEventListener("click", handleRunAI);
saveGroqKeyBtn.addEventListener("click", handleSaveGroqKey);

// Carregar chave salva
chrome.storage.sync.get(["groq_api_key"], (result) => {
  if (result.groq_api_key) {
    groqKeyInput.value = result.groq_api_key;
  }
});

function handleAuth() {
  authBtn.disabled = true;
  sheetsList.innerHTML = '<div class="loading">Carregando planilhas...</div>';

  getToken()
    .then((token) => listSheets(token))
    .then((sheets) => {
      currentSheets = sheets;
      displaySheets(sheets);
    })
    .catch((error) => {
      console.error(error);
      sheetsList.innerHTML = `<div class="error">Erro: ${error.message}</div>`;
    })
    .finally(() => {
      authBtn.disabled = false;
    });
}

function handleSheetClick(e) {
  const li = e.target.closest("li");
  if (!li) return;

  const index = Array.from(sheetsList.children).indexOf(li);
  const sheet = currentSheets[index];
  openSheetPanel(sheet);
}

function handleBack() {
  sheetPanel.style.display = "none";
  mainPanel.style.display = "block";
}

function handleSaveGroqKey() {
  const key = groqKeyInput.value.trim();
  chrome.storage.sync.set({ groq_api_key: key }, () => {
    sheetResult.textContent = key
      ? "Chave salva com sucesso!"
      : "Chave removida.";
    sheetResult.className = "success";
  });
}

function handleRunAI() {
  if (!selectedSheet) return;

  const groqKey = groqKeyInput.value.trim();
  if (!groqKey) {
    sheetResult.textContent = "Por favor, adicione sua chave do Groq primeiro.";
    sheetResult.className = "error";
    return;
  }

  runSheetAI.disabled = true;
  sheetResult.textContent = "Conectando com Groq AI...";
  sheetResult.className = "";

  processWithGroq(selectedSheet, sheetPrompt.value, groqKey)
    .then((result) => {
      sheetResult.textContent = result;
      sheetResult.className = "";
    })
    .catch((error) => {
      console.error(error);
      sheetResult.textContent = `Erro: ${error.message}`;
      sheetResult.className = "error";
    })
    .finally(() => {
      runSheetAI.disabled = false;
    });
}

function getToken() {
  return new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive: true }, (token) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
      } else {
        resolve(token);
      }
    });
  });
}

async function listSheets(token) {
  const response = await fetch(
    "https://www.googleapis.com/drive/v3/files?q=mimeType='application/vnd.google-apps.spreadsheet'&fields=files(id,name,createdTime,modifiedTime)&orderBy=modifiedTime desc",
    {
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
    }
  );

  if (!response.ok) {
    const errorData = await response.json();
    throw new Error(errorData.error?.message || "Erro ao buscar planilhas");
  }

  const data = await response.json();
  return data.files || [];
}

function displaySheets(sheets) {
  sheetsList.innerHTML = "";

  if (sheets.length === 0) {
    const li = document.createElement("li");
    li.textContent = "Nenhuma planilha encontrada";
    sheetsList.appendChild(li);
    return;
  }

  sheets.forEach((sheet) => {
    const li = document.createElement("li");
    const sheetName = document.createElement("div");
    sheetName.className = "sheet-name";
    sheetName.textContent = sheet.name;

    const sheetDetails = document.createElement("div");
    sheetDetails.className = "sheet-details";
    sheetDetails.textContent = `Criada em: ${new Date(
      sheet.createdTime
    ).toLocaleDateString("pt-BR")}`;

    li.appendChild(sheetName);
    li.appendChild(sheetDetails);
    sheetsList.appendChild(li);
  });
}

function openSheetPanel(sheet) {
  selectedSheet = sheet;
  sheetTitle.textContent = sheet.name;
  sheetLink.href = `https://docs.google.com/spreadsheets/d/${sheet.id}`;
  sheetPanel.style.display = "block";
  mainPanel.style.display = "none";
  sheetResult.textContent = "A resposta aparecerá aqui...";
  sheetResult.className = "";
}

async function processWithGroq(sheet, userPrompt, apiKey) {
  try {
    sheetResult.textContent = "Obtendo dados da planilha...";

    const token = await getToken();
    const title = await getFirstSheetTitle(sheet.id, token);
    const range = buildA1RangeForFirstPage(title);
    const values = await getSheetValues(sheet.id, range, token);
    const tsv = valuesToTSV(values);
    const usedRange = computeUsedRangeA1(title, values);
    const prompt = buildCommandPrompt({
      sheetTitle: title,
      usedRange,
      userText: userPrompt,
      sheetTSV: tsv,
    });

    sheetResult.textContent = "Processando com Groq AI...";
    const response = await callGroqAPI(prompt, apiKey);

    const parsed = tryParseJson(response);
    if (parsed) {
      return JSON.stringify(parsed, null, 2);
    }
    return response;
  } catch (error) {
    console.error("Erro no processamento:", error);
    throw new Error(`Falha ao processar: ${error.message}`);
  }
}

async function getFirstSheetTitle(spreadsheetId, token) {
  const resp = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets(properties(title,index))`,
    {
      headers: { Authorization: `Bearer ${token}` },
    }
  );

  if (!resp.ok) throw new Error("Erro ao obter abas da planilha");

  const data = await resp.json();
  const firstSheet = (data.sheets || [])[0];
  if (!firstSheet) throw new Error("Planilha sem abas");

  return firstSheet.properties.title;
}

async function getSheetValues(spreadsheetId, range, token) {
  const resp = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
      range
    )}`,
    {
      headers: { Authorization: `Bearer ${token}` },
    }
  );

  if (!resp.ok) throw new Error("Erro ao obter dados da planilha");

  const data = await resp.json();
  return data.values || [];
}

function buildA1RangeForFirstPage(sheetTitle) {
  return `'${sheetTitle.replace(/'/g, "''")}'!A1:Z1000`;
}

function valuesToTSV(values) {
  if (!values || values.length === 0) return "";
  return values
    .map((row) => row.map((cell) => String(cell ?? "")).join("\t"))
    .join("\n");
}

function buildCommandPrompt({ sheetTitle, usedRange, userText, sheetTSV }) {
  const schema = `
RETORNE APENAS JSON, sem comentários ou Markdown.
Estrutura do JSON:
{
  "plan": string, // resumo breve do que será feito
  "commands": [
    {
      "type": "transform_values", // preferencial para operações aritméticas
      "range": string, // A1 da área alvo, ex: '${usedRange}'
      "onlyNumeric": boolean, // true para afetar apenas números
      "expression": string // expressão usando x como valor da célula, ex: "x*2"
    }
    // Alternativas (quando necessário):
    // { "type": "set_formula", "range": string, "formulaTemplate": string }
    // { "type": "write_values", "range": string, "values": number[][] | string[][] }
  ]
}`.trim();

  return [
    "Você é um gerador de comandos para Google Sheets.",
    "Sua saída deve ser exclusivamente um JSON válido seguindo o schema descrito.",
    "Não explique, não use Markdown, não inclua texto fora do JSON.",
    "",
    `TÍTULO DA ABA: ${sheetTitle}`,
    `INTERVALO UTILIZADO (estimado): ${usedRange}`,
    "",
    "DADOS (TSV, primeira aba):",
    sheetTSV || "(vazio)",
    "",
    "PEDIDO DO USUÁRIO:",
    userText || "",
    "",
    "SCHEMA DE RESPOSTA:",
    schema,
  ].join("\n");
}

function tryParseJson(text) {
  if (!text) return null;
  // Remove cercas de código se vierem
  const fenced = text.match(/```(?:json)?\n([\s\S]*?)\n```/);
  const raw = fenced ? fenced[1] : text;
  // Tenta extrair o maior bloco JSON
  const match = raw.match(/\{[\s\S]*\}/);
  const candidate = match ? match[0] : raw;
  try {
    return JSON.parse(candidate);
  } catch (_) {
    return null;
  }
}

function computeUsedRangeA1(sheetTitle, values) {
  const numRows = values?.length || 0;
  let numCols = 0;
  for (const row of values || []) {
    if (Array.isArray(row)) numCols = Math.max(numCols, row.length);
  }
  if (numRows === 0 || numCols === 0)
    return `'${sheetTitle.replace(/'/g, "''")}'!A1:A1`;
  const endCol = toA1Col(numCols);
  return `'${sheetTitle.replace(/'/g, "''")}'!A1:${endCol}${numRows}`;
}

function toA1Col(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

async function callGroqAPI(prompt, apiKey) {
  const body = {
    model: "llama3-70b-8192",
    messages: [
      {
        role: "system",
        content:
          "Você é um assistente especializado em análise de dados e planilhas. Responda em português de forma clara e objetiva.",
      },
      { role: "user", content: prompt },
    ],
    temperature: 0.3,
    max_tokens: 2000,
    stream: false,
  };

  const response = await fetch(
    "https://api.groq.com/openai/v1/chat/completions",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    const errorData = await response.json();
    throw new Error(errorData.error?.message || "Erro na API Groq");
  }

  const data = await response.json();
  return (
    data.choices[0]?.message?.content || "Não foi possível obter resposta."
  );
}
