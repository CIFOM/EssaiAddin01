/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Boutons
    document.getElementById("depart-ex01").onclick = () => {
      document.getElementById("confirm-ex01").style.display = "block";
    };
    document.getElementById("confirm-yes").onclick = departEx01;
    document.getElementById("confirm-no").onclick = () => {
      document.getElementById("confirm-ex01").style.display = "none";
    };
    document.getElementById("verifier-ex01").onclick = verifierEx01;
  }
});

// Fonction Départ Ex01
export async function departEx01() {
  // Fermer la confirmation
  document.getElementById("confirm-ex01").style.display = "none";
  // Remet la croix rouge au départ
  document.getElementById("verif-status").textContent = "❌";

  return Word.run(async (context) => {
    context.document.body.clear();

    context.document.body.insertParagraph("Ceci doit être un titre de niveau 1", Word.InsertLocation.end);
    context.document.body.insertParagraph("Ceci doit être un paragraphe normal", Word.InsertLocation.end);
    context.document.body.insertParagraph("Ceci doit être un titre de niveau 2", Word.InsertLocation.end);
    context.document.body.insertParagraph("Ceci doit être un paragraphe normal", Word.InsertLocation.end);
    context.document.body.insertParagraph("Ceci doit être un autre titre de niveau 2", Word.InsertLocation.end);
    context.document.body.insertParagraph("Ceci doit être un paragraphe normal: appliquez les niveaux de titre indiqués et pressez \"Vérifier Ex01\"", Word.InsertLocation.end);

    await context.sync();
  });
}

// Vérification des styles
export async function verifierEx01() {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/style");
    await context.sync();

    // Styles attendus Bureau
    const attenduBureau = [
      "Titre 1", // 1
      "Normal",    // 2
      "Titre 2", // 3
      "Normal",    // 4
      "Titre 2", // 5
      "Normal"     // 6
    ];
    // Styles attendus Web
    const attenduWeb = [
      "Heading 1", // Ligne 1
      "Normal",    // Ligne 2
      "Heading 2", // Ligne 3
      "Normal",    // Ligne 4
      "Heading 2", // Ligne 5
      "Normal"     // Ligne 6
    ];;

    let ok = true;
    for (let i = 0; i < 6; i++) {
      if (!paragraphs.items[i] || (paragraphs.items[i].style !== attenduBureau[i] && paragraphs.items[i].style !== attenduWeb[i])) {
        ok = false;
        break;
      }
    }

    // Mise à jour de l’icône
    document.getElementById("verif-status").textContent = ok ? "✅" : "❌";
  });
}
