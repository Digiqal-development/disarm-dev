/* global Word */

import { taggedObjects } from "./object-tag-component.js";
export let centralNode = null;
let selectedCentralNode = null;

export async function insertObjectSummaryTables() {
    console.log("=== insertObjectSummaryTables CALLED ===");
    await Word.run(async (context) => {
        const body = context.document.body;

        const centralTypes = ["Threat Actor", "Campaign", "Incident"];

        const centralCandidates = [];

        centralTypes.forEach(type => {
            if (taggedObjects[type]) {
                taggedObjects[type].forEach(obj => {
                    centralCandidates.push ({type, value: obj.value});
                });
            }
        });

        if (centralCandidates.length === 1) {
            console.log("AUTO CENTRAL NODE:", centralCandidates[0]);
        }
        console.log("CHECK popup condition", {
            count: centralCandidates.length,
            selectedCentralNode
        });
        if (centralCandidates.length > 1 && !selectedCentralNode) {
            openCentralNodePicker(centralCandidates);
            return;
        }

        const hasCentralNode = centralTypes.some(
            (type) => taggedObjects[type] && taggedObjects[type].length > 0
        );

        if (!hasCentralNode) {
            document.getElementById("main").insertAdjacentHTML(
                "afterbegin",
                "<p style='color:red'>Please tag or add at least one Threat Actor, Campaign, or Incident.</p>"
            );
            return;
        }

        const selectedTypes = getSelectedObjectTypes();
        const typesToInsert =
            selectedTypes.length > 0
                ? selectedTypes
                : Object.keys(taggedObjects);
        console.log("selectedTypes:", selectedTypes);
        typesToInsert.forEach((objectType) => {
            const items = taggedObjects[objectType];
            console.log("ITEMS:", items);
            if (!items || items.length === 0) return;

            const table = body.insertTable(
                items.length + 1,
                3,
                Word.InsertLocation.end
            );

            table.getCell(0, 0).value = objectType;
            table.getCell(0, 1).value = "Use";
            table.getCell(0, 2).value = "Occurrences";

            items.forEach((item, index) => {
                table.getCell(index + 1, 0).value = item.value;
                table.getCell(index + 1, 1).value = item.use;
                table.getCell(index + 1, 2).value = item.occurrences.toString();
            });
        });

        await context.sync();
    });
}

//open central node picker function
function openCentralNodePicker(candidates) {
    selectedCentralNode = null;

    const popup = document.getElementById("popup-central-node");
    const table = document.getElementById("central-node-table");
    const selectBtn = document.getElementById("select-central-node-btn");
    const closeBtn = document.getElementById("close-central-node");

    table.innerHTML = "";

    candidates.forEach((c) => {
        const tr = document.createElement("tr");
        tr.classList.add("list", "option1");
        tr.innerHTML = `<td>${c.value} <span style="opacity:.7">[${c.type}]</span></td>`;

        tr.onclick = () => {
            Array.from(table.querySelectorAll("tr")).forEach(r => r.classList.remove("selected"));
            tr.classList.add("selected");
            selectedCentralNode = c;
            centralNode = c;
        };

        table.appendChild(tr);
    });

    // wire buttons
    selectBtn.onclick = () => {
        if (!selectedCentralNode) return;
        centralNode = selectedCentralNode;
        selectedCentralNode = centralNode;
        console.log("central node set:", centralNode);
        popup.style.display = "none";
        insertObjectSummaryTables();
    };
    closeBtn.onclick = () => popup.style.display = "none";

    popup.style.display = "block";
}


function closeCentralNodePicker() {
    const popup = document.getElementById("popup-central-node");
    popup.style.display = "none";
}

function getSelectedObjectTypes() {
    return Array.from(
        document.querySelectorAll("#object-summary-selection input[type=checkbox]:checked")
    ).map(cb => cb.value);
}

