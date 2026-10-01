// GCP Construction - Accueil SST
// Génère le PDF dans le navigateur puis l'envoie au Cloudflare Worker,
// qui l'enregistre dans SharePoint.

const WORKER_URL = "https://gcp-sst-sharepoint.philip-brunet.workers.dev/";
const DOWNLOAD_LOCAL_COPY = false;

const form = document.getElementById("sstForm");
const statusEl = document.getElementById("status");
const canvas = document.getElementById("signatureCanvas");
const resetBtn = document.getElementById("resetBtn");

let signaturePad;


// ======================================================
// STATUT
// ======================================================

function setStatus(message, type = "") {
  if (!statusEl) return;

  statusEl.textContent = message;
  statusEl.className = `status ${type}`;
}


// ======================================================
// SIGNATURE
// ======================================================

function initSignature() {
  if (!canvas || typeof SignaturePad === "undefined") return;

  signaturePad = new SignaturePad(canvas, {
    backgroundColor: "rgb(255,255,255)",
    penColor: "rgb(0,0,0)"
  });

  resizeSignatureCanvas();
}


function resizeSignatureCanvas() {
  if (!canvas || !signaturePad) return;

  const ratio = Math.max(window.devicePixelRatio || 1, 1);

  canvas.width = canvas.offsetWidth * ratio;
  canvas.height = canvas.offsetHeight * ratio;

  canvas.getContext("2d").scale(ratio, ratio);

  signaturePad.clear();
}


window.addEventListener("resize", resizeSignatureCanvas);


// ======================================================
// OUTILS
// ======================================================

function formatDateFr(dateValue) {
  if (!dateValue) return "-";

  const parts = dateValue.split("-");

  if (parts.length !== 3) return dateValue;

  return `${parts[2]}-${parts[1]}-${parts[0]}`;
}


function sanitizeFileName(value) {
  return String(value || "")
    .replace(/[\\/:*?"<>|]/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}


function fileToDataURL(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;

    reader.readAsDataURL(file);
  });
}


async function compressImage(file, maxDimension = 1400, quality = 0.82) {
  const source = await fileToDataURL(file);

  return new Promise((resolve, reject) => {
    const img = new Image();

    img.onload = () => {
      let width = img.width;
      let height = img.height;

      if (width > height && width > maxDimension) {
        height = Math.round(height * maxDimension / width);
        width = maxDimension;
      }

      if (height >= width && height > maxDimension) {
        width = Math.round(width * maxDimension / height);
        height = maxDimension;
      }

      const tempCanvas = document.createElement("canvas");

      tempCanvas.width = width;
      tempCanvas.height = height;

      const ctx = tempCanvas.getContext("2d");

      ctx.drawImage(img, 0, 0, width, height);

      resolve(
        tempCanvas.toDataURL(
          "image/jpeg",
          quality
        )
      );
    };

    img.onerror = reject;

    img.src = source;
  });
}


async function loadLogoAsDataUrl() {
  try {
    const response = await fetch("logo.png", {
      cache: "no-store"
    });

    if (!response.ok) return null;

    const blob = await response.blob();

    return await fileToDataURL(blob);

  } catch (error) {
    console.warn("Logo non disponible :", error);

    return null;
  }
}


function addWrappedText(
  doc,
  text,
  x,
  y,
  maxWidth,
  lineHeight = 4.5
) {
  const lines = doc.splitTextToSize(text || "-", maxWidth);

  doc.text(lines, x, y);

  return y + lines.length * lineHeight;
}


function addField(
  doc,
  label,
  value,
  x,
  y,
  maxWidth
) {
  doc.setFont("helvetica", "bold");
  doc.text(label, x, y);

  const labelWidth =
    doc.getTextWidth(label) + 3;

  doc.setFont("helvetica", "normal");

  const lines = doc.splitTextToSize(
    value || "-",
    maxWidth - labelWidth
  );

  doc.text(
    lines,
    x + labelWidth,
    y
  );

  return y + Math.max(1, lines.length) * 6;
}


// ======================================================
// RÉCUPÉRATION DU FORMULAIRE
// ======================================================

async function getFormData() {
  const selectedSubjects = Array.from(
    document.querySelectorAll(
      'input[name="sujets"]:checked'
    )
  ).map(input => input.value);


  const photoInput =
    document.getElementById("photoCompetence");


  let photoCompetence = null;


  if (
    photoInput &&
    photoInput.files &&
    photoInput.files[0]
  ) {
    photoCompetence =
      await compressImage(
        photoInput.files[0]
      );
  }


  const autreSujetCheck =
    document.getElementById(
      "autreSujetCheck"
    );


  const autreSujet =
    document.getElementById(
      "autreSujet"
    );


  return {
    chantier:
      document.getElementById("chantier")?.value.trim() || "",

    entreprise:
      document.getElementById("entreprise")?.value.trim() || "",

    travailleur:
      document.getElementById("travailleur")?.value.trim() || "",

    telephone:
      document.getElementById("telephone")?.value.trim() || "",

    dateArrivee:
      document.getElementById("dateArrivee")?.value || "",

    metier:
      document.getElementById("metier")?.value.trim() || "",

    sujets:
      selectedSubjects,

    autreSujetCheck:
      !!autreSujetCheck?.checked,

    autreSujet:
      autreSujet?.value.trim() || "",

    photoCompetence,

    signatureDataUrl:
      signaturePad && !signaturePad.isEmpty()
        ? signaturePad.toDataURL("image/png")
        : null
  };
}


// ======================================================
// CRÉATION DU PDF
// ======================================================

async function createPdf(data) {
  const { jsPDF } = window.jspdf;

  const doc = new jsPDF({
    unit: "mm",
    format: "letter",
    orientation: "portrait"
  });

  const pageWidth =
    doc.internal.pageSize.getWidth();

  const pageHeight =
    doc.internal.pageSize.getHeight();

  const margin = 14;

  const contentWidth =
    pageWidth - margin * 2;

  let y = 9;


  // ====================================================
  // EN-TÊTE
  // ====================================================

  const logoData =
    await loadLogoAsDataUrl();


  if (logoData) {
    try {
      doc.addImage(
        logoData,
        "PNG",
        margin,
        y,
        39,
        15
      );
    } catch (error) {
      console.warn(
        "Impossible d'ajouter le logo",
        error
      );
    }
  }


  doc.setTextColor(15, 42, 75);

  doc.setFont(
    "helvetica",
    "bold"
  );

  doc.setFontSize(15);


  doc.text(
    "ACCUEIL DES TRAVAILLEURS",
    pageWidth - margin,
    y + 7,
    {
      align: "right"
    }
  );


  doc.setFont(
    "helvetica",
    "normal"
  );

  doc.setFontSize(8.5);

  doc.setTextColor(
    70,
    80,
    90
  );


  doc.text(
    "GCP Construction - Accueil SST",
    pageWidth - margin,
    y + 13,
    {
      align: "right"
    }
  );


  y += 20;


  doc.setDrawColor(
    225,
    165,
    40
  );

  doc.setLineWidth(0.8);


  doc.line(
    margin,
    y,
    pageWidth - margin,
    y
  );


  y += 7;


  // ====================================================
  // 1. IDENTIFICATION
  // ====================================================

  doc.setTextColor(
    15,
    42,
    75
  );

  doc.setFont(
    "helvetica",
    "bold"
  );

  doc.setFontSize(11);


  doc.text(
    "1. Identification",
    margin,
    y
  );


  y += 7;


  doc.setTextColor(
    0,
    0,
    0
  );

  doc.setFontSize(8.7);


  const colGap = 8;

  const colWidth =
    (contentWidth - colGap) / 2;


  const leftX =
    margin;

  const rightX =
    margin + colWidth + colGap;


  const row1 =
    y;


  addField(
    doc,
    "Chantier :",
    data.chantier,
    leftX,
    row1,
    colWidth
  );


  addField(
    doc,
    "Entreprise :",
    data.entreprise,
    rightX,
    row1,
    colWidth
  );


  addField(
    doc,
    "Travailleur :",
    data.travailleur,
    leftX,
    row1 + 6,
    colWidth
  );


  addField(
    doc,
    "Téléphone :",
    data.telephone,
    rightX,
    row1 + 6,
    colWidth
  );


  addField(
    doc,
    "Date d'arrivée :",
    formatDateFr(
      data.dateArrivee
    ),
    leftX,
    row1 + 12,
    colWidth
  );


  addField(
    doc,
    "Métier :",
    data.metier || "-",
    rightX,
    row1 + 12,
    colWidth
  );


  y += 23;


  // ====================================================
  // 2. SUJETS DISCUTÉS
  // ====================================================

  doc.setTextColor(
    15,
    42,
    75
  );

  doc.setFont(
    "helvetica",
    "bold"
  );

  doc.setFontSize(11);


  doc.text(
    "2. Sujets discutés",
    margin,
    y
  );


  y += 6;


  doc.setTextColor(
    0,
    0,
    0
  );

  doc.setFont(
    "helvetica",
    "normal"
  );

  doc.setFontSize(8.5);


  const subjects =
    [...data.sujets];


  if (
    data.autreSujetCheck &&
    data.autreSujet
  ) {
    subjects.push(
      `Autre sujet : ${data.autreSujet}`
    );
  }


  if (subjects.length) {
    const half =
      Math.ceil(
        subjects.length / 2
      );


    const columns = [
      subjects.slice(
        0,
        half
      ),

      subjects.slice(
        half
      )
    ];


    const subjectStartY =
      y;


    let maxRows = 0;


    columns.forEach(
      (items, colIndex) => {

        items.forEach(
          (subject, i) => {

            const x =
              colIndex === 0
                ? leftX
                : rightX;


            doc.text(
              "•",
              x,
              subjectStartY + i * 5
            );


            const short =
              doc.splitTextToSize(
                subject,
                colWidth - 5
              );


            doc.text(
              short,
              x + 4,
              subjectStartY + i * 5
            );
          }
        );


        maxRows =
          Math.max(
            maxRows,
            items.length
          );
      }
    );


    y +=
      maxRows * 5 + 4;

  } else {
    doc.text(
      "-",
      margin,
      y
    );

    y += 7;
  }


  // ====================================================
  // 3. CARTE + 4. ENGAGEMENT
  // ====================================================

  doc.setTextColor(
    15,
    42,
    75
  );

  doc.setFont(
    "helvetica",
    "bold"
  );

  doc.setFontSize(11);


  doc.text(
    "3. Carte de compétence",
    margin,
    y
  );


  doc.text(
    "4. Engagement SST",
    rightX,
    y
  );


  y += 6;


  const blockTop =
    y;


  const photoW =
    colWidth;


  const photoH =
    62;


  // ====================================================
  // PHOTO
  // ====================================================

  doc.setDrawColor(
    210,
    215,
    220
  );


  doc.rect(
    leftX,
    blockTop,
    photoW,
    photoH
  );


  if (data.photoCompetence) {
    try {
      const props =
        doc.getImageProperties(
          data.photoCompetence
        );


      const ratio =
        props.width /
        props.height;


      let drawW =
        photoW - 2;


      let drawH =
        drawW / ratio;


      if (
        drawH >
        photoH - 2
      ) {
        drawH =
          photoH - 2;

        drawW =
          drawH * ratio;
      }


      const drawX =
        leftX +
        (photoW - drawW) / 2;


      const drawY =
        blockTop +
        (photoH - drawH) / 2;


      doc.addImage(
        data.photoCompetence,
        "JPEG",
        drawX,
        drawY,
        drawW,
        drawH
      );


    } catch (error) {
      console.warn(
        "Photo non intégrée",
        error
      );


      doc.setFont(
        "helvetica",
        "normal"
      );


      doc.setFontSize(8);


      doc.setTextColor(
        0,
        0,
        0
      );


      doc.text(
        "Photo fournie - aperçu indisponible.",
        leftX + 2,
        blockTop + 6
      );
    }

  } else {
    doc.setFont(
      "helvetica",
      "normal"
    );


    doc.setFontSize(8);


    doc.setTextColor(
      100,
      100,
      100
    );


    doc.text(
      "Aucune photo fournie.",
      leftX + 3,
      blockTop + 7
    );
  }


  // ====================================================
  // ENGAGEMENT
  // ====================================================

  doc.setTextColor(
    0,
    0,
    0
  );


  doc.setFont(
    "helvetica",
    "normal"
  );


  doc.setFontSize(8.2);


  const engagement =
    "Je m'engage à respecter la planification sécuritaire du maître d'œuvre ainsi que les parties du programme de prévention qui me sont applicables. Je comprends qu'un manquement pourrait entraîner des mesures disciplinaires ou des sanctions.";


  let engagementY =
    blockTop;


  engagementY =
    addWrappedText(
      doc,
      engagement,
      rightX,
      engagementY,
      colWidth,
      4.3
    );


  engagementY += 5;


  // ====================================================
  // SIGNATURE
  // ====================================================

  doc.setFont(
    "helvetica",
    "bold"
  );


  doc.setFontSize(8.5);


  doc.text(
    "Signature du travailleur :",
    rightX,
    engagementY
  );


  engagementY += 3;


  doc.setDrawColor(
    220,
    220,
    220
  );


  doc.rect(
    rightX,
    engagementY,
    colWidth,
    28
  );


  if (data.signatureDataUrl) {
    try {
      doc.addImage(
        data.signatureDataUrl,
        "PNG",
        rightX + 2,
        engagementY + 2,
        colWidth - 4,
        24
      );

    } catch (error) {
      console.warn(
        "Signature non intégrée",
        error
      );
    }
  }


  // ====================================================
  // ATTESTATION
  // ====================================================

  y =
    blockTop +
    photoH +
    8;


  doc.setDrawColor(
    225,
    165,
    40
  );


  doc.setLineWidth(0.5);


  doc.line(
    margin,
    y,
    pageWidth - margin,
    y
  );


  y += 6;


  doc.setFont(
    "helvetica",
    "normal"
  );


  doc.setFontSize(7.8);


  doc.setTextColor(
    70,
    70,
    70
  );


  doc.text(
    "Le travailleur confirme avoir pris connaissance des éléments ci-dessus et accepte l'engagement SST.",
    margin,
    y
  );


  // ====================================================
  // PIED DE PAGE
  // ====================================================

  doc.setFontSize(7.3);


  doc.setTextColor(
    110,
    110,
    110
  );


  doc.text(
    `Document généré le ${new Date().toLocaleString("fr-CA")}`,
    margin,
    pageHeight - 8
  );


  doc.text(
    "GCP Construction",
    pageWidth - margin,
    pageHeight - 8,
    {
      align: "right"
    }
  );


  return doc;
}


// ======================================================
// ENVOI VERS SHAREPOINT + TURNSTILE
// ======================================================

async function uploadPdfToSharePoint(
  pdfBlob,
  fileName
) {

  const payload =
    new FormData();


  // ====================================================
  // RÉCUPÉRER LE JETON TURNSTILE
  // ====================================================

  const turnstileToken =
    document.querySelector(
      'input[name="cf-turnstile-response"]'
    )?.value;


  if (!turnstileToken) {
    throw new Error(
      "Veuillez compléter la vérification de sécurité."
    );
  }


  // ====================================================
  // CONTENU À TRANSMETTRE
  // ====================================================

  payload.append(
    "pdf",
    pdfBlob,
    fileName
  );


  payload.append(
    "fileName",
    fileName
  );


  payload.append(
    "turnstileToken",
    turnstileToken
  );


  // ====================================================
  // ENVOI AU WORKER
  // ====================================================

  const response =
    await fetch(
      WORKER_URL,
      {
        method: "POST",
        body: payload
      }
    );


  let result = null;


  try {
    result =
      await response.json();

  } catch {
    result = null;
  }


  if (!response.ok) {
    const details =
      result?.details ||
      result?.error ||
      result?.message ||
      `Erreur HTTP ${response.status}`;


    throw new Error(
      details
    );
  }


  if (
    result &&
    result.ok === false
  ) {
    throw new Error(
      result.error ||
      result.message ||
      "Erreur lors de l'envoi."
    );
  }


  return result;
}


// ======================================================
// SOUMISSION
// ======================================================

form?.addEventListener(
  "submit",
  async event => {

    event.preventDefault();


    try {
      setStatus(
        "Préparation de l'accueil SST...",
        "loading"
      );


      // ====================================================
      // VÉRIFIER LA SIGNATURE
      // ====================================================

      if (
        signaturePad &&
        signaturePad.isEmpty()
      ) {
        setStatus(
          "Veuillez signer l'accueil SST avant de transmettre.",
          "error"
        );

        return;
      }


      // ====================================================
      // VÉRIFIER TURNSTILE AVANT LE PDF
      // ====================================================

      const turnstileToken =
        document.querySelector(
          'input[name="cf-turnstile-response"]'
        )?.value;


      if (!turnstileToken) {
        setStatus(
          "Veuillez compléter la vérification de sécurité.",
          "error"
        );

        return;
      }


      const data =
        await getFormData();


      setStatus(
        "Génération du PDF...",
        "loading"
      );


      const doc =
        await createPdf(data);


      const safeDate =
        sanitizeFileName(
          data.dateArrivee ||
          new Date()
            .toISOString()
            .slice(0, 10)
        );


      const safeWorker =
        sanitizeFileName(
          data.travailleur ||
          "Travailleur"
        );


      const safeSite =
        sanitizeFileName(
          data.chantier ||
          "Chantier"
        );


      const fileName =
        `${safeDate} - ${safeWorker} - ${safeSite} - Accueil SST.pdf`;


      const pdfBlob =
        doc.output("blob");


      setStatus(
        "Transmission vers SharePoint...",
        "loading"
      );


      const result =
        await uploadPdfToSharePoint(
          pdfBlob,
          fileName
        );


      if (DOWNLOAD_LOCAL_COPY) {
        doc.save(
          fileName
        );
      }


      setStatus(
        `Accueil transmis avec succès dans SharePoint : ${fileName}`,
        "success"
      );


      const submitButton =
        form.querySelector(
          'button[type="submit"]'
        );


      if (submitButton) {
        submitButton.dataset.originalText =
          submitButton.dataset.originalText ||
          submitButton.textContent;


        submitButton.textContent =
          "✓ Transmis";
      }


      // ====================================================
      // RÉINITIALISER TURNSTILE APRÈS SUCCÈS
      // ====================================================

      if (
        typeof turnstile !== "undefined"
      ) {
        try {
          turnstile.reset();
        } catch (error) {
          console.warn(
            "Turnstile reset impossible :",
            error
          );
        }
      }


      console.log(
        "SharePoint :",
        result
      );


    } catch (error) {
      console.error(
        error
      );


      setStatus(
        `Erreur : ${error.message}`,
        "error"
      );


      // Réinitialise Turnstile après une erreur
      if (
        typeof turnstile !== "undefined"
      ) {
        try {
          turnstile.reset();
        } catch (resetError) {
          console.warn(
            "Turnstile reset impossible :",
            resetError
          );
        }
      }
    }
  }
);


// ======================================================
// RÉINITIALISATION
// ======================================================

resetBtn?.addEventListener(
  "click",
  () => {

    form?.reset();


    if (signaturePad) {
      signaturePad.clear();
    }


    // Réinitialiser Turnstile
    if (
      typeof turnstile !== "undefined"
    ) {
      try {
        turnstile.reset();
      } catch (error) {
        console.warn(
          "Turnstile reset impossible :",
          error
        );
      }
    }


    setStatus(
      "",
      ""
    );


    const submitButton =
      form?.querySelector(
        'button[type="submit"]'
      );


    if (
      submitButton &&
      submitButton.dataset.originalText
    ) {
      submitButton.textContent =
        submitButton.dataset.originalText;
    }
  }
);


// ======================================================
// DÉMARRAGE
// ======================================================

window.addEventListener(
  "load",
  () => {
    initSignature();
  }
);
