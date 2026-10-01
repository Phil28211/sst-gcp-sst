// GCP Construction - Accueil SST
// Génère le PDF dans le navigateur puis l'envoie au Cloudflare Worker,
// qui l'enregistre dans SharePoint.

const WORKER_URL = "https://gcp-sst-sharepoint.philip-brunet.workers.dev/";
const DOWNLOAD_LOCAL_COPY = false; // mettre true si vous voulez aussi télécharger le PDF localement

const form = document.getElementById("sstForm");
const statusEl = document.getElementById("status");
const canvas = document.getElementById("signatureCanvas");
const resetBtn = document.getElementById("resetBtn");

let signaturePad;

function setStatus(message, type = "") {
  statusEl.textContent = message;
  statusEl.className = `status ${type}`;
}

function initSignature() {
  signaturePad = new SignaturePad(canvas, {
    backgroundColor: "rgb(255,255,255)",
    penColor: "rgb(0,0,0)"
  });

  resizeSignatureCanvas();
}

function resizeSignatureCanvas() {
  if (!canvas || !signaturePad) return;

  const ratio = Math.max(window.devicePixelRatio || 1, 1);
  const rect = canvas.getBoundingClientRect();

  const previousData = !signaturePad.isEmpty() ? signaturePad.toData() : null;

  canvas.width = rect.width * ratio;
  canvas.height = rect.height * ratio;

  const ctx = canvas.getContext("2d");
  ctx.scale(ratio, ratio);

  signaturePad.clear();

  if (previousData) {
    signaturePad.fromData(previousData);
  }
}

function checkedValues(name) {
  return Array.from(
    document.querySelectorAll(`input[name="${name}"]:checked`)
  ).map(el => el.value);
}

function sanitizeFileName(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[<>:"/\\|?*]/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}

function formatDateForFile(dateValue) {
  if (!dateValue) {
    const d = new Date();
    return d.toISOString().slice(0, 10);
  }
  return dateValue;
}

function formatDateFr(dateValue) {
  if (!dateValue) return "-";
  const [year, month, day] = dateValue.split("-");
  return `${day}-${month}-${year}`;
}

async function imageFileToCompressedDataUrl(fileInput, maxDimension = 1400, quality = 0.82) {
  const file = fileInput.files && fileInput.files[0];
  if (!file) return null;

  const objectUrl = URL.createObjectURL(file);

  try {
    const img = await new Promise((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = reject;
      image.src = objectUrl;
    });

    let width = img.naturalWidth;
    let height = img.naturalHeight;

    const largest = Math.max(width, height);
    if (largest > maxDimension) {
      const scale = maxDimension / largest;
      width = Math.round(width * scale);
      height = Math.round(height * scale);
    }

    const tempCanvas = document.createElement("canvas");
    tempCanvas.width = width;
    tempCanvas.height = height;

    const ctx = tempCanvas.getContext("2d");
    ctx.drawImage(img, 0, 0, width, height);

    return tempCanvas.toDataURL("image/jpeg", quality);
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

async function loadLogoAsDataUrl() {
  try {
    const response = await fetch("logo.png", { cache: "no-store" });
    if (!response.ok) return null;

    const blob = await response.blob();

    return await new Promise(resolve => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => resolve(null);
      reader.readAsDataURL(blob);
    });
  } catch {
    return null;
  }
}

function addWrappedText(doc, text, x, y, maxWidth, lineHeight = 5.5) {
  const lines = doc.splitTextToSize(text || "-", maxWidth);
  doc.text(lines, x, y);
  return y + lines.length * lineHeight;
}

function addField(doc, label, value, x, y, maxWidth) {
  doc.setFont("helvetica", "bold");
  doc.text(label, x, y);

  doc.setFont("helvetica", "normal");
  const labelWidth = doc.getTextWidth(label) + 3;
  const lines = doc.splitTextToSize(value || "-", maxWidth - labelWidth);
  doc.text(lines, x + labelWidth, y);

  return y + Math.max(1, lines.length) * 6;
}

async function createPdf(data) {
  const { jsPDF } = window.jspdf;

  const doc = new jsPDF({
    unit: "mm",
    format: "letter",
    orientation: "portrait"
  });

  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 15;
  const contentWidth = pageWidth - margin * 2;

  let y = 14;

  // Logo
  const logoData = await loadLogoAsDataUrl();
  if (logoData) {
    try {
      doc.addImage(logoData, "PNG", (pageWidth - 48) / 2, y, 48, 22);
      y += 27;
    } catch {
      y += 2;
    }
  }

  doc.setTextColor(15, 42, 75);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(17);
  doc.text("ACCUEIL DES TRAVAILLEURS", pageWidth / 2, y, { align: "center" });

  y += 7;
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.setTextColor(70, 80, 90);
  doc.text("GCP Construction - Accueil SST", pageWidth / 2, y, { align: "center" });

  y += 10;
  doc.setDrawColor(225, 165, 40);
  doc.setLineWidth(1);
  doc.line(margin, y, pageWidth - margin, y);
  y += 10;

  // Identification
  doc.setTextColor(15, 42, 75);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.text("1. Identification", margin, y);

  y += 9;
  doc.setTextColor(0, 0, 0);
  doc.setFontSize(10);

  y = addField(doc, "Nom du chantier :", data.chantier, margin, y, contentWidth);
  y = addField(doc, "Nom de l'entreprise :", data.entreprise, margin, y, contentWidth);
  y = addField(doc, "Nom du travailleur :", data.travailleur, margin, y, contentWidth);
  y = addField(doc, "Téléphone :", data.telephone, margin, y, contentWidth);
  y = addField(doc, "Date d'arrivée :", formatDateFr(data.dateArrivee), margin, y, contentWidth);
  y = addField(doc, "Métier :", data.metier || "-", margin, y, contentWidth);

  y += 4;

  // Sujets discutés
  doc.setTextColor(15, 42, 75);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.text("2. Sujets discutés", margin, y);

  y += 8;
  doc.setTextColor(0, 0, 0);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);

  const subjects = [...data.sujets];
  if (data.autreSujetCheck && data.autreSujet) {
    subjects.push(`Autre sujet : ${data.autreSujet}`);
  }

  if (subjects.length === 0) {
    doc.text("-", margin, y);
    y += 6;
  } else {
    subjects.forEach(subject => {
      doc.text("•", margin, y);
      y = addWrappedText(doc, subject, margin + 5, y, contentWidth - 5, 5.5);
      y += 1;
    });
  }

  y += 4;

  // Carte de compétence
  doc.setTextColor(15, 42, 75);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.text("3. Carte de compétence", margin, y);
  y += 8;

  if (data.photoCompetence) {
    // Si la photo ne tient pas bien sur la page actuelle, nouvelle page
    if (y > pageHeight - 95) {
      doc.addPage();
      y = 15;
    }

    try {
      doc.addImage(data.photoCompetence, "JPEG", margin, y, 85, 64);
      y += 70;
    } catch {
      doc.setFont("helvetica", "normal");
      doc.setTextColor(0, 0, 0);
      doc.text("Photo fournie, mais impossible de l'intégrer au PDF.", margin, y);
      y += 7;
    }
  } else {
    doc.setFont("helvetica", "normal");
    doc.setTextColor(0, 0, 0);
    doc.text("Aucune photo fournie.", margin, y);
    y += 8;
  }

  // Engagement
  if (y > pageHeight - 80) {
    doc.addPage();
    y = 15;
  }

  doc.setTextColor(15, 42, 75);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.text("4. Engagement SST", margin, y);
  y += 8;

  doc.setTextColor(0, 0, 0);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9.5);

  const engagement =
    "Je m'engage à respecter la planification sécuritaire du maître d'œuvre ainsi que les parties du programme de prévention qui me sont applicables. Je comprends qu'un manquement pourrait entraîner des mesures disciplinaires ou des sanctions.";

  y = addWrappedText(doc, engagement, margin, y, contentWidth, 5);

  y += 7;
  doc.setFont("helvetica", "bold");
  doc.text("Signature du travailleur :", margin, y);
  y += 4;

  if (data.signatureDataUrl) {
    try {
      doc.addImage(data.signatureDataUrl, "PNG", margin, y, 70, 27);
      y += 32;
    } catch {
      y += 4;
    }
  }

  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.5);
  doc.setTextColor(90, 90, 90);
  doc.text(
    `Document généré le ${new Date().toLocaleString("fr-CA")}`,
    margin,
    Math.min(y + 4, pageHeight - 10)
  );

  return doc;
}

async function uploadPdfToSharePoint(pdfBlob, fileName) {
  const formData = new FormData();
  formData.append("pdf", pdfBlob, fileName);
  formData.append("fileName", fileName);

  const response = await fetch(WORKER_URL, {
    method: "POST",
    body: formData
  });

  let result;
  try {
    result = await response.json();
  } catch {
    throw new Error(`Réponse invalide du serveur (HTTP ${response.status}).`);
  }

  if (!response.ok || !result.ok) {
    const detail =
      result?.details?.error?.message ||
      result?.details?.error_description ||
      result?.error ||
      `Erreur HTTP ${response.status}`;

    throw new Error(detail);
  }

  return result;
}

form.addEventListener("submit", async event => {
  event.preventDefault();

  if (!form.checkValidity()) {
    form.reportValidity();
    return;
  }

  if (!signaturePad || signaturePad.isEmpty()) {
    setStatus("Veuillez signer le formulaire avant de l'envoyer.", "error");
    return;
  }

  const submitButton = form.querySelector('button[type="submit"]');
  const originalButtonText = submitButton.textContent;

  try {
    submitButton.disabled = true;
    submitButton.textContent = "Transmission en cours...";
    setStatus("Préparation du PDF...", "");

    const photoCompetence = await imageFileToCompressedDataUrl(
      document.getElementById("photoCompetence")
    );

    const data = {
      chantier: document.getElementById("chantier").value.trim(),
      entreprise: document.getElementById("entreprise").value.trim(),
      travailleur: document.getElementById("travailleur").value.trim(),
      telephone: document.getElementById("telephone").value.trim(),
      dateArrivee: document.getElementById("dateArrivee").value,
      metier: document.getElementById("metier").value.trim(),
      sujets: checkedValues("sujets"),
      autreSujetCheck: document.getElementById("autreSujetCheck").checked,
      autreSujet: document.getElementById("autreSujet").value.trim(),
      photoCompetence,
      signatureDataUrl: signaturePad.toDataURL("image/png")
    };

    const pdfDoc = await createPdf(data);

    const fileName =
      `${formatDateForFile(data.dateArrivee)} - ` +
      `${sanitizeFileName(data.travailleur)} - ` +
      `${sanitizeFileName(data.chantier)} - Accueil SST.pdf`;

    const pdfBlob = pdfDoc.output("blob");

    setStatus("Envoi du PDF vers SharePoint...", "");

    const result = await uploadPdfToSharePoint(pdfBlob, fileName);

    if (DOWNLOAD_LOCAL_COPY) {
      pdfDoc.save(fileName);
    }

    setStatus(
      `Accueil transmis avec succès dans SharePoint : ${result.fileName || fileName}`,
      "success"
    );

    submitButton.textContent = "✓ Transmis";

  } catch (error) {
    console.error(error);
    setStatus(
      `Échec de la transmission : ${error.message}`,
      "error"
    );
    submitButton.textContent = originalButtonText;
  } finally {
    submitButton.disabled = false;
  }
});

resetBtn.addEventListener("click", () => {
  form.reset();

  if (signaturePad) {
    signaturePad.clear();
  }

  setStatus("", "");

  const submitButton = form.querySelector('button[type="submit"]');
  submitButton.disabled = false;
  submitButton.textContent = "▧ Générer le PDF";
});

window.addEventListener("resize", () => {
  window.clearTimeout(window.__gcpResizeTimer);
  window.__gcpResizeTimer = window.setTimeout(resizeSignatureCanvas, 150);
});

document.addEventListener("DOMContentLoaded", initSignature);
