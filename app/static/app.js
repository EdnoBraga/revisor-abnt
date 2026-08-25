const form = document.querySelector("#revision-form");
const fileInput = document.querySelector("#document");
const fileName = document.querySelector("#file-name");
const result = document.querySelector("#result");
const submitButton = document.querySelector("#submit-button");

fileInput.addEventListener("change", () => {
  fileName.textContent = fileInput.files[0]?.name || "Selecione o arquivo";
});

function showResult(message, kind = "progress") {
  result.hidden = false;
  result.className = `result ${kind}`;
  result.innerHTML = message;
}

async function poll(statusUrl) {
  const response = await fetch(statusUrl);
  const job = await response.json();
  if (!response.ok) throw new Error(job.detail || "Não foi possível consultar a revisão.");
  if (job.status === "completed") {
    showResult(`
      <h2>Revisão concluída</h2>
      <p>Baixe o novo DOCX e o relatório. Abra o arquivo no Word para atualizar os campos do sumário e conferir a paginação.</p>
      <p><a class="button-link" href="${job.download_url}">Baixar DOCX revisado</a><a class="text-link" href="${job.audit_url}">Baixar relatório</a></p>
    `, "success");
    submitButton.disabled = false;
    submitButton.textContent = "Revisar outro arquivo";
    return;
  }
  if (job.status === "failed") {
    showResult(`<h2>Não foi possível concluir</h2><p>${job.error}</p>`, "error");
    submitButton.disabled = false;
    submitButton.textContent = "Tentar novamente";
    return;
  }
  setTimeout(() => poll(statusUrl).catch((error) => showResult(`<p>${error.message}</p>`, "error")), 1200);
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!fileInput.files[0]) return;
  submitButton.disabled = true;
  submitButton.textContent = "Enviando…";
  showResult("<h2>Arquivo recebido</h2><p>Estamos verificando a estrutura e aplicando a formatação.</p>");
  const data = new FormData(form);
  try {
    const response = await fetch("/api/revisions", { method: "POST", body: data });
    const job = await response.json();
    if (!response.ok) throw new Error(job.detail || "O envio não foi aceito.");
    submitButton.textContent = "Revisando…";
    poll(job.status_url).catch((error) => showResult(`<p>${error.message}</p>`, "error"));
  } catch (error) {
    showResult(`<h2>Envio não realizado</h2><p>${error.message}</p>`, "error");
    submitButton.disabled = false;
    submitButton.textContent = "Revisar arquivo";
  }
});
