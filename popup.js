const authBtn = document.getElementById("auth");
const sheetsList = document.getElementById("sheets");

const sheetPanel = document.getElementById("sheetPanel");
const sheetTitle = document.getElementById("sheetTitle");
const sheetLink = document.getElementById("sheetLink");
const sheetPrompt = document.getElementById("sheetPrompt");
const sheetResult = document.getElementById("sheetResult");
const backBtn = document.getElementById("backBtn");
const runSheetAI = document.getElementById("runSheetAI");

sheetsList.addEventListener("click", (e) => {
  let li = e.target.closest("li");
  if (!li) return;

  const index = Array.from(sheetsList.children).indexOf(li);
  const sheet = currentSheets[index]; // precisamos guardar as sheets carregadas
  openSheetPanel(sheet);
});

let currentSheets = []; // guarda planilhas carregadas

authBtn.addEventListener("click", async () => {
  try {
    authBtn.disabled = true;
    sheetsList.innerHTML = '<div class="loading">Carregando planilhas...</div>';

    const token = await getToken();
    const sheets = await listSheets(token);
    currentSheets = sheets; // salvar para clique

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
  } catch (err) {
    console.error(err);
    sheetsList.innerHTML = `<div class="error">Erro ao buscar planilhas: ${err.message}</div>`;
  } finally {
    authBtn.disabled = false;
  }
});

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

function openSheetPanel(sheet) {
  sheetTitle.textContent = sheet.name;
  sheetLink.href = `https://docs.google.com/spreadsheets/d/${sheet.id}`;
  sheetPanel.style.display = "block";
  sheetsList.parentElement.style.display = "none"; // esconde lista
  sheetResult.textContent = "";
}

backBtn.addEventListener("click", () => {
  sheetPanel.style.display = "none";
  sheetsList.parentElement.style.display = "block"; // volta para lista
});
