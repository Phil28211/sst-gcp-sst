const EMAIL_ENABLED=false;
const EMAILJS_PUBLIC_KEY="REMPLACER_PAR_PUBLIC_KEY";
const EMAILJS_SERVICE_ID="REMPLACER_PAR_SERVICE_ID";
const EMAILJS_TEMPLATE_ID="REMPLACER_PAR_TEMPLATE_ID";
const EMAIL_TO="sst@gcpconstruction.ca";

const form=document.getElementById("sstForm");
const statusEl=document.getElementById("status");
const canvas=document.getElementById("signatureCanvas");
const clearSignatureBtn=document.getElementById("clearSignature");
let signaturePad;

function setStatus(m,t=""){statusEl.textContent=m;statusEl.className=`status ${t}`;}
function resizeCanvas(){const ratio=Math.max(window.devicePixelRatio||1,1);const rect=canvas.getBoundingClientRect();const data=signaturePad&&!signaturePad.isEmpty()?signaturePad.toData():null;canvas.width=rect.width*ratio;canvas.height=rect.height*ratio;canvas.getContext("2d").scale(ratio,ratio);if(signaturePad){signaturePad.clear();if(data)signaturePad.fromData(data);}}
function initSignature(){signaturePad=new SignaturePad(canvas,{backgroundColor:"rgb(255,255,255)",penColor:"rgb(0,0,0)"});resizeCanvas();}
window.addEventListener("resize",resizeCanvas);
clearSignatureBtn.addEventListener("click",()=>signaturePad.clear());

function checked(name){return Array.from(document.querySelectorAll(`input[name="${name}"]:checked`)).map(x=>x.value);}
function clean(v){return String(v||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^a-zA-Z0-9-_ ]/g,"").replace(/\s+/g," ").trim();}
function readFile(input){return new Promise(resolve=>{const file=input.files&&input.files[0];if(!file)return resolve(null);const r=new FileReader();r.onload=()=>resolve({name:file.name,dataUrl:r.result});r.onerror=()=>resolve(null);r.readAsDataURL(file);});}
function wrap(doc,text,x,y,w,lh=6){const lines=doc.splitTextToSize(text||"-",w);doc.text(lines,x,y);return y+lines.length*lh;}
function field(doc,label,value,x,y){doc.setFont("helvetica","bold");doc.text(label,x,y);doc.setFont("helvetica","normal");doc.text(value||"-",x+55,y);return y+8;}

async function createPdf(data){
 const {jsPDF}=window.jspdf;const doc=new jsPDF({unit:"mm",format:"letter"});const W=doc.internal.pageSize.getWidth();const m=14;let y=16;
 doc.setFillColor(31,58,95);doc.rect(0,0,W,22,"F");doc.setTextColor(255,255,255);doc.setFont("helvetica","bold");doc.setFontSize(16);doc.text("ACCUEIL SST - GCP CONSTRUCTION",m,14);
 doc.setTextColor(0,0,0);doc.setFontSize(11);y=34;doc.setFont("helvetica","bold");doc.text("1. Identification",m,y);y+=9;
 y=field(doc,"Nom du chantier :",data.chantier,m,y);y=field(doc,"Entreprise :",data.entreprise,m,y);y=field(doc,"Travailleur :",data.travailleur,m,y);y=field(doc,"Téléphone :",data.telephone,m,y);y=field(doc,"Date d'arrivée :",data.dateArrivee,m,y);y=field(doc,"Métier :",data.metier,m,y);y=field(doc,"Syndicat :",data.syndicat,m,y);y=field(doc,"No certificat :",data.certificat,m,y);
 y+=4;doc.setFont("helvetica","bold");doc.text("2. Sujets discutés",m,y);y+=8;doc.setFont("helvetica","normal");y=wrap(doc,data.sujets.join(", ")||"-",m,y,W-m*2);if(data.autreSujet){y+=2;y=wrap(doc,"Autre : "+data.autreSujet,m,y,W-m*2);}
 y+=6;doc.setFont("helvetica","bold");doc.text("3. Formations et compétences",m,y);y+=8;doc.setFont("helvetica","normal");y=wrap(doc,data.formations.join(", ")||"-",m,y,W-m*2);
 y+=6;doc.setFont("helvetica","bold");doc.text("4. Engagement SST",m,y);y+=8;doc.setFont("helvetica","normal");y=wrap(doc,"Le travailleur confirme avoir pris connaissance des consignes SST, du programme de prévention applicable et s'engage à respecter les règles du chantier.",m,y,W-m*2);
 y+=10;doc.setFont("helvetica","bold");doc.text("Signature du travailleur :",m,y);y+=4;if(data.signatureDataUrl){doc.addImage(data.signatureDataUrl,"PNG",m,y,70,28);y+=34;}doc.setFont("helvetica","normal");doc.text(`Généré le : ${new Date().toLocaleString("fr-CA")}`,m,y);
 if(data.photoCompetence||data.photoAsp){doc.addPage();y=16;doc.setFont("helvetica","bold");doc.setFontSize(14);doc.text("Photos des cartes et certifications",m,y);y+=12;
  if(data.photoCompetence){doc.setFontSize(11);doc.text("Carte de compétence / carte principale",m,y);y+=4;try{doc.addImage(data.photoCompetence.dataUrl,"JPEG",m,y,85,65);}catch{try{doc.addImage(data.photoCompetence.dataUrl,"PNG",m,y,85,65);}catch{}}y+=75;}
  if(data.photoAsp){doc.setFontSize(11);doc.text("Carte ASP / formation supplémentaire",m,y);y+=4;try{doc.addImage(data.photoAsp.dataUrl,"JPEG",m,y,85,65);}catch{try{doc.addImage(data.photoAsp.dataUrl,"PNG",m,y,85,65);}catch{}}}
 }
 return doc;
}
async function sendEmailWithPdf(pdfDoc,data,fileName){
 if(!EMAIL_ENABLED)return false;emailjs.init({publicKey:EMAILJS_PUBLIC_KEY});
 const pdfBase64=pdfDoc.output("datauristring").split(",")[1];
 await emailjs.send(EMAILJS_SERVICE_ID,EMAILJS_TEMPLATE_ID,{to_email:EMAIL_TO,travailleur:data.travailleur,chantier:data.chantier,entreprise:data.entreprise,date_arrivee:data.dateArrivee,message:"Un nouvel accueil SST a été complété.",pdf_attachment:pdfBase64,pdf_filename:fileName});
 return true;
}
form.addEventListener("submit",async e=>{
 e.preventDefault();if(signaturePad.isEmpty()){setStatus("Veuillez signer avant de générer le PDF.","error");return;}
 try{setStatus("Génération du PDF en cours...");
 const data={chantier:document.getElementById("chantier").value.trim(),entreprise:document.getElementById("entreprise").value.trim(),travailleur:document.getElementById("travailleur").value.trim(),telephone:document.getElementById("telephone").value.trim(),dateArrivee:document.getElementById("dateArrivee").value,metier:document.getElementById("metier").value.trim(),syndicat:document.getElementById("syndicat").value.trim(),certificat:document.getElementById("certificat").value.trim(),autreSujet:document.getElementById("autreSujet").value.trim(),sujets:checked("sujets"),formations:checked("formations"),signatureDataUrl:signaturePad.toDataURL("image/png"),photoCompetence:await readFile(document.getElementById("photoCompetence")),photoAsp:await readFile(document.getElementById("photoAsp"))};
 const pdfDoc=await createPdf(data);const fileName=`${clean(data.travailleur)} - ${clean(data.chantier)} - Accueil SST.pdf`;pdfDoc.save(fileName);
 if(EMAIL_ENABLED){setStatus("PDF généré. Envoi courriel en cours...");await sendEmailWithPdf(pdfDoc,data,fileName);setStatus("PDF généré et courriel envoyé avec succès.","success");}else{setStatus("PDF généré avec succès. Le téléchargement devrait être lancé.","success");}
 }catch(err){console.error(err);setStatus("Une erreur est survenue pendant la génération du PDF.","error");}
});
document.addEventListener("DOMContentLoaded",initSignature);
