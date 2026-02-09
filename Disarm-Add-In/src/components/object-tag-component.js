/*global Word*/

import * as helpersUI from "../utils/helpers-ui.js";

export const taggedObjects = {};

//popup open
export function insertObjectTag() {
    helpersUI.blurBackground();

    const rows = document.querySelectorAll("#object-type-table tr");

    //clear previous selection
    rows.forEach(tr => tr.classList.remove("selected"));

    //single select behaviour
    rows.forEach(row => {
        row.onclick = (e) => {
            e.stopPropagation();

            rows.forEach(r => r.classList.remove("selected"));
            row.classList.add("selected");
        };
    });

    document.getElementById("popup-insert-object-tag").style.display = "block";
}

// save tag
export async function saveObjectTag() {
    await Word.run(async (context) => {
        const selectedRow = document.querySelector("#object-type-table tr.selected");
        if (!selectedRow) return;
        const objectType = selectedRow.textContent.trim();

        const selection = context.document.getSelection();
        const paragraphs = selection.paragraphs;

        selection.load("text");
        paragraphs.load("items/text");
        await context.sync();

        console.log("DEBUG selection.text =", selection.text);
        console.log("DEBUG paragraphs =", paragraphs.items.length);

        if (paragraphs.items.length === 0) return;

        const firstPara = paragraphs.items[0];
        firstPara.load("text");
        await context.sync();
        const raw = selection.text.trim();
        const paraText = firstPara.text;

        const idx = paraText.indexOf(raw);

        let useText = "";

        if (idx !== -1) {
            useText = paraText.substring(idx).trim();
        } else {
            useText = paraText.trim();
        }

        try {
            const firstPara = paragraphs.items[0];
            const paraRange = firstPara.getRange();

            const tables = context.document.body.tables;
            tables.load("items");
            await context.sync();

            let foundIndex = -1;

            for (let i = 0; i < tables.items.length; i++) {

                const rel = paraRange.compareLocationWith(tables.items[i].getRange());
                await context.sync();

                if (rel.value === Word.LocationRelation.inside) {
                    foundIndex = i + 1;
                    break;
                }
            }

            if (foundIndex !== -1) {
                useText = `Table ${foundIndex}`;
                console.log("TABLE INDEX FOUND:", useText);
            }

        } catch (e) {
            console.log("Table detection failed:", e.message);
        }

        const parts = raw.split(",").map(p => p.trim()).filter(Boolean);

        const taggedText = parts
            .map(p => `${p} [${objectType}]`)
            .join(", ");

        if (!taggedObjects[objectType]) taggedObjects[objectType] = [];

        parts.forEach((val) => {
            const existing = taggedObjects[objectType].find(o => o.value === val);

            if (existing) {
                existing.occurrences++;
                existing.use = useText;
            } else {
                taggedObjects[objectType].push({
                    value: val,
                    use: useText,
                    occurrences: 1,
                });
            }
        });

        console.log("ABOUT TO INSERT:", taggedText);
        console.log("Using:", useText);

        selection.insertText(taggedText, Word.InsertLocation.replace);
        await context.sync();

        // highlight only the [ObjectType] part
        const para = context.document.getSelection().paragraphs.getFirst();
        para.load("text");
        await context.sync();

        const search = para.search(`[${objectType}]`);
        search.load("items");
        await context.sync();

        search.items.forEach(r => {
            r.font.set({
                highlightColor: "#FFFF00"
            });
        });

        helpersUI.closeAllFields();
        helpersUI.unblurBackground();
    }).catch(err => {
        console.error("Save Error:", err.message);
    });
}



//function for add object manually
export async function addObjectManually() {
    const nameInput = document.getElementById("new-object-name");
    const typeSelect = document.getElementById("new-object-type");

    const objectName = nameInput.value.trim();
    const objectType = typeSelect.value;

    if (!taggedObjects[objectType]) {
        taggedObjects[objectType] = [];
    }

    taggedObjects[objectType].push({
        value: objectName,
        use: "Added manually",
        occurrences: 1,
    });

    //reset fields
    nameInput.value = "";

    //close popup
    document.getElementById("popup-add-object").style.display = "none";
}

async function getTableCaption(table, context) {
    table.load("isNullObject,rows");
    await context.sync();

    if (table.isNullObject) return "Table";

    const before = table.getRange(Word.RangeLocation.before);
    const paragraphs = before.paragraphs;
    paragraphs.load("items/text");
    await context.sync();

    const paras = paragraphs.items;
    for (let i = paras.length - 1; i >= 0; i--) {
        const txt = paras[i].text.trim().toLowerCase();
        if (txt.startsWith("table") || txt.includes("tablica")) {
            return paras[i].text.trim();
        }
    }
    return "Table";  // Fallback
}

//popup close
document.getElementById("close-add-object")?.addEventListener("click", (e) => {
    e.stopPropagation();

    document.getElementById("popup-add-object").style.display = "none";
});

// when closing main object popup → force close child popup too
document.getElementById("close-object")?.addEventListener("click", () => {
    document.getElementById("popup-add-object").style.display = "none";
});




