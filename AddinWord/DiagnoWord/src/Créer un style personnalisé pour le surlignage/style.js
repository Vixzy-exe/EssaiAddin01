/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("run").onclick = run;
  }
});

async function run() {
  return Word.run(async (context) => {
    const body = context.document.body;

    // Vérification si le document est vide
    if (body.text.trim().length > 0) {
      const confirmation = confirm("Le document contient déjà du contenu. Voulez-vous l'effacer ?");
      if (!confirmation) {
        return;
      }
      body.clear(); // Efface le contenu existant
    }

    // Création d’un document avec du texte simple
    body.insertParagraph("Instructions de l'exercice :", Word.InsertLocation.start);
    body.insertParagraph("1. Créez un style personnalisé.", Word.InsertLocation.end);
    body.insertParagraph("2. Appliquez ce style à un mot ou un paragraphe désigné.", Word.InsertLocation.end);
    body.insertParagraph("3. Vérifiez que le style est correctement appliqué.", Word.InsertLocation.end);

    // Création d’un style personnalisé
    const customStyleName = "StylePersonnalise";
    const styles = context.document.styles;
    styles.add(customStyleName, Word.StyleType.paragraph);
    const customStyle = styles.getByName(customStyleName);
    customStyle.font.color = "red";
    customStyle.font.bold = true;
    customStyle.font.italic = true;
    customStyle.font.highlightColor = "yellow";

    // Application du style personnalisé
    const paragraph = body.insertParagraph("Texte formaté avec le style personnalisé.", Word.InsertLocation.end);
    paragraph.styleBuiltIn = customStyleName;

    // Synchronisation avec Word
    await context.sync();

    // Validation
    const appliedStyle = paragraph.styleBuiltIn;
    if (appliedStyle === customStyleName) {
      console.log("Le style personnalisé a été correctement appliqué.");
    } else {
      console.error("Le style personnalisé n'a pas été appliqué correctement.");
    }
  });
}