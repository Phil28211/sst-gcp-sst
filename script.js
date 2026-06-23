const form = document.getElementById("sstForm");
const statusEl = document.getElementById("status");
const canvas = document.getElementById("signatureCanvas");
const signatureBox = document.querySelector(".signature-box");
const resetBtn = document.getElementById("resetBtn");
let signaturePad;

function setStatus(message, type = "") {
  statusEl.textContent = message;
  statusEl.className = `status ${type}`;
}

function resizeCanvas() {
  const ratio = Math.max(window.devicePixelRatio || 1, 1);
  const rect = canvas.getBoundingClientRect();
  const existing = signaturePad && !signaturePad.isEmpty() ? signaturePad.toData() : null;

  canvas.width = rect.width * ratio;
  canvas.height = rect.height * ratio;
  canvas.getContext("2d").scale(ratio, ratio);

  if (signaturePad) {
    signaturePad.clear();
    if (existing) signaturePad.fromData(existing);
  }
}

function initSignature() {
  signaturePad = new SignaturePad(canvas, {
    backgroundColor: "rgb(255,255,255)",
    penColor: "rgb(0,0,0)"
  });
  signaturePad.addEventListener("beginStroke", () => signatureBox.classList.add("has-signature"));
  resizeCanvas();
}

window.addEventListener("resize", resizeCanvas);
window.addEventListener("orientationchange", () => setTimeout(resizeCanvas, 250));

document.addEventListener("DOMContentLoaded", initSignature);

resetBtn.addEventListener("click", () => {
  form.reset();
  signaturePad.clear();
  signatureBox.classList.remove("has-signature");
  setStatus("");
});

function checkedValues(name) {
  return Array.from(document.querySelectorAll(`input[name="${name}"]:checked`)).map(item => item.value);
}

function sanitizeFileName(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9-_ ]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function readFileAsDataUrl(input) {
  return new Promise(resolve => {
    const file = input.files && input.files[0];
    if (!file) return resolve(null);

    const reader = new FileReader();
    reader.onload = () => resolve({ name: file.name, dataUrl: reader.result });
    reader.onerror = () => resolve(null);
    reader.readAsDataURL(file);
  });
}

function addField(doc, label, value, x, y) {
  doc.setFont("helvetica", "bold");
  doc.text(label, x, y);
  doc.setFont("helvetica", "normal");
  doc.text(value || "-", x + 55, y);
  return y + 8;
}

function addWrappedText(doc, text, x, y, maxWidth, lineHeight = 6) {
  const lines = doc.splitTextToSize(text || "-", maxWidth);
  doc.text(lines, x, y);
  return y + lines.length * lineHeight;
}

async function createPdf(data) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "mm", format: "letter" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const margin = 14;
  let y = 16;

  doc.setFillColor(8, 38, 75);
  doc.rect(0, 0, pageWidth, 22, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  doc.text("ACCUEIL DES TRAVAILLEURS - GCP CONSTRUCTION", margin, 14);

  doc.setTextColor(0, 0, 0);
  doc.setFontSize(11);
  y = 34;

  doc.setFont("helvetica", "bold");
  doc.text("1. Identification", margin, y);
  y += 9;
  y = addField(doc, "Nom du chantier :", data.chantier, margin, y);
  y = addField(doc, "Entreprise :", data.entreprise, margin, y);
  y = addField(doc, "Travailleur :", data.travailleur, margin, y);
  y = addField(doc, "Téléphone :", data.telephone, margin, y);
  y = addField(doc, "Date d'arrivée :", data.dateArrivee, margin, y);
  y = addField(doc, "Métier :", data.metier, margin, y);

  y += 6;
  doc.setFont("helvetica", "bold");
  doc.text("2. Sujets discutés", margin, y);
  y += 8;
  doc.setFont("helvetica", "normal");
  const sujets = [...data.sujets];
  if (data.autreSujetCheck && data.autreSujet) sujets.push("Autre sujet : " + data.autreSujet);
  y = addWrappedText(doc, sujets.join(", ") || "-", margin, y, pageWidth - margin * 2);

  y += 6;
  doc.setFont("helvetica", "bold");
  doc.text("3. Formations et cartes", margin, y);
  y += 8;
  doc.setFont("helvetica", "normal");
  doc.text(data.photoCompetence ? "Photo de carte de compétence jointe au PDF." : "Aucune photo de carte jointe.", margin, y);

  y += 12;
  doc.setFont("helvetica", "bold");
  doc.text("4. Engagement SST", margin, y);
  y += 8;
  doc.setFont("helvetica", "normal");
  y = addWrappedText(doc, "Je m'engage à respecter la planification sécuritaire du maître d'œuvre ainsi que les parties du programme de prévention qui me sont applicables. Je comprends qu'un manquement pourrait entraîner des mesures disciplinaires ou des sanctions.", margin, y, pageWidth - margin * 2);

  y += 10;
  doc.setFont("helvetica", "bold");
  doc.text("Signature du travailleur :", margin, y);
  y += 4;
  doc.addImage(data.signatureDataUrl, "PNG", margin, y, 75, 30);

  if (data.photoCompetence) {
    doc.addPage();
    y = 16;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(14);
    doc.text("Photo carte de compétence", margin, y);
    y += 8;
    try {
      doc.addImage(data.photoCompetence.dataUrl, "JPEG", margin, y, 120, 90);
    } catch {
      try { doc.addImage(data.photoCompetence.dataUrl, "PNG", margin, y, 120, 90); } catch {}
    }
  }

  return doc;
}

form.addEventListener("submit", async event => {
  event.preventDefault();

  if (signaturePad.isEmpty()) {
    setStatus("Veuillez signer avant de générer le PDF.", "error");
    return;
  }

  try {
    setStatus("Génération du PDF en cours...");

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
      photoCompetence: await readFileAsDataUrl(document.getElementById("photoCompetence")),
      signatureDataUrl: signaturePad.toDataURL("image/png")
    };

    const pdfDoc = await createPdf(data);
    const fileName = `${sanitizeFileName(data.travailleur)} - ${sanitizeFileName(data.chantier)} - Accueil SST.pdf`;
    pdfDoc.save(fileName);
    setStatus("PDF généré avec succès.", "success");
  } catch (error) {
    console.error(error);
    setStatus("Une erreur est survenue pendant la génération du PDF.", "error");
  }
});
