Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("createDocument").onclick = createDocument;
    document.getElementById("resetForm").onclick = resetForm;

    ["gender", "title", "firstName", "lastName"].forEach((fieldId) => {
      document.getElementById(fieldId).addEventListener("input", updatePreview);
      document.getElementById(fieldId).addEventListener("change", updatePreview);
    });

    updatePreview();
  }
});

async function createDocument() {
  const gender = document.getElementById("gender").value;
  const title = document.getElementById("title").value.trim();
  const firstName = document.getElementById("firstName").value.trim();
  const lastName = document.getElementById("lastName").value.trim();
  const statusMessage = document.getElementById("statusMessage");
  const button = document.getElementById("createDocument");

  clearStatus();
  clearFieldErrors();

  if (!gender) {
    markFieldError("gender");
    showStatus("Bitte Geschlecht auswählen.", "error");
    return;
  }

  if (!firstName) {
    markFieldError("firstName");
    showStatus("Bitte Vorname eingeben.", "error");
    return;
  }

  if (!lastName) {
    markFieldError("lastName");
    showStatus("Bitte Nachname eingeben.", "error");
    return;
  }

  const today = new Date().toLocaleDateString("de-CH", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  });

  const salutation = gender === "frau" ? "Liebe Frau" : "Lieber Herr";
  const belegarzt = gender === "frau" ? "Belegärztin" : "Belegarzt";

const formattedTitle = title ? `${title} ` : "";

const replacements = {
  "[DATUM]": today,
  "[ANREDE]": salutation,
  "[TITEL]": formattedTitle,
  "[VORNAME]": firstName,
  "[NACHNAME]": lastName,
  "[BELEGARZT]": belegarzt,
};

  try {
    button.disabled = true;
    showStatus("Dokument wird erstellt …", "info");

    await Word.run(async (context) => {
    let replacementCount = 0;

    for (const [placeholder, value] of Object.entries(replacements)) {
      replacementCount += await replaceInRange(context.document.body, placeholder, value);

      const sections = context.document.sections;
      sections.load("items");
      await context.sync();

      for (const section of sections.items) {
        const header = section.getHeader("Primary");
        const footer = section.getFooter("Primary");

        replacementCount += await replaceInRange(header, placeholder, value);
        replacementCount += await replaceInRange(footer, placeholder, value);
      }
    }

    if (replacementCount === 0) {
      throw new Error("Keine Platzhalter im Dokument gefunden.");
    }
    });

    showStatus("Dokument wurde erfolgreich erstellt.", "success");
    button.textContent = "Dokument erstellt";
    button.disabled = true;
  } catch (error) {
    console.error(error);

    if (error.message === "Keine Platzhalter im Dokument gefunden.") {
      showStatus("Es wurden keine Platzhalter im Dokument gefunden.", "error");
    } else {
      showStatus("Fehler beim Erstellen des Dokuments. Details siehe Konsole.", "error");
    }
  } finally {
    if (button.textContent !== "Dokument erstellt") {
      button.disabled = false;
    }
  }
}

function showStatus(message, type) {
  const statusMessage = document.getElementById("statusMessage");
  statusMessage.textContent = message;
  statusMessage.className = type ? `status ${type}` : "status";
}

function clearStatus() {
  const statusMessage = document.getElementById("statusMessage");
  statusMessage.textContent = "";
  statusMessage.className = "status";
}

function markFieldError(fieldId) {
  const field = document.getElementById(fieldId);
  field.classList.add("field-error");
  field.focus();
}

function clearFieldErrors() {
  ["gender", "firstName", "lastName"].forEach((fieldId) => {
    document.getElementById(fieldId).classList.remove("field-error");
  });
}


function resetForm() {
  document.getElementById("gender").value = "";
  document.getElementById("title").value = "";
  document.getElementById("firstName").value = "";
  document.getElementById("lastName").value = "";

  clearStatus();
  clearFieldErrors();

  const button = document.getElementById("createDocument");
  button.textContent = "Dokument erstellen";
  button.disabled = false;

  updatePreview();
}


async function replaceInRange(range, placeholder, value) {
  const searchResults = range.search(placeholder, {
    matchCase: true,
    matchWholeWord: false,
  });

  searchResults.load("items");
  await range.context.sync();

  const count = searchResults.items.length;

  searchResults.items.forEach((result) => {
    result.insertText(value, Word.InsertLocation.replace);
  });

  await range.context.sync();

  return count;
}

function updatePreview() {
  const gender = document.getElementById("gender").value;
  const title = document.getElementById("title").value.trim();
  const firstName = document.getElementById("firstName").value.trim();
  const lastName = document.getElementById("lastName").value.trim();

  const today = new Date().toLocaleDateString("de-CH", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  });

  const salutation = gender === "frau" ? "Liebe Frau" : gender === "mann" ? "Lieber Herr" : "";
  const role = gender === "frau" ? "Belegärztin" : gender === "mann" ? "Belegarzt" : "";

  const formattedTitle = title ? `${title} ` : "";
  const fullName = `${formattedTitle}${lastName}`.trim();
  const firstAndLastName = `${firstName} ${lastName}`.trim();

  document.getElementById("previewSalutation").textContent =
    salutation || fullName ? `${salutation} ${fullName}`.trim() : "Bitte Personendaten eingeben.";

  document.getElementById("previewDetails").textContent =
    firstAndLastName || "–";

  document.getElementById("previewRole").textContent =
    role || "–";

  document.getElementById("previewDate").textContent =
    today;
}


