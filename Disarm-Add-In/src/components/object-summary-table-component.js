/* global Word */

import { taggedObjects } from "./object-tag-component.js";
export let centralNode = null;
let selectedCentralNode = null;

export async function insertObjectSummaryTables() {
    console.log("=== insertObjectSummaryTables CALLED ===");
    await Word.run(async (context) => {
        const body = context.document.body;

        // clear memory
        Object.keys(taggedObjects).forEach(k => delete taggedObjects[k]);

        body.load("text");
        await context.sync();

        const fullText = body.text;
        const tagRegex = /(\b[^\s]+)\s*\[([A-Za-z ]+)\]/g;

        let match;

        while ((match = tagRegex.exec(fullText)) !== null) {
            const value = match[1].trim();
            const type = match[2].trim();

            if (!taggedObjects[type]) taggedObjects[type] = [];

            const existing = taggedObjects[type].find(o => o.value === value);

            if (existing) {
                existing.occurrences++;
            } else {
                taggedObjects[type].push({
                    value,
                    use: extractSentenceContaining(fullText, value),
                    occurrences: 1
                });
            }
        }

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
            centralNode = centralCandidates[0];
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
        await sendGraphToBackend();
    });
}

//open central node picker function
function openCentralNodePicker(candidates) {
    selectedCentralNode = null;

    const popup = document.getElementById("popup-central-node");
    const table = document.getElementById("central-node-table");
    const tableBody = table.tBodies[0] || table.appendChild(document.createElement("tbody"));

    tableBody.innerHTML = "";

    candidates.forEach((c, index) => {
        const tr = document.createElement("tr");
        tr.classList.add("list", "option1");
        tr.dataset.index = index;
        tr.innerHTML = `<td>${c.value} <span style="opacity:.7">[${c.type}]</span></td>`;

        tr.addEventListener('click', function (e) {
            Array.from(tableBody.querySelectorAll("tr")).forEach(r => {
                r.classList.remove("selected");
            });

            this.classList.add("selected");
            this.style.backgroundColor = '#55ace3';
            this.style.color = 'white';
            Array.from(this.cells).forEach(cell => {
                cell.style.backgroundColor = '#55ace3';
                cell.style.color = 'white';
            });
            selectedCentralNode = c;
            centralNode = c;
            console.log("Selected:", c);
        });

        tr.addEventListener('mouseenter', function () {
            if (!this.classList.contains('selected')) {
                this.style.backgroundColor = '#f5f5f5';
            }
        });

        tr.addEventListener('mouseleave', function () {
            if (!this.classList.contains('selected')) {
                this.style.backgroundColor = '';
            }
        });

        tableBody.appendChild(tr);
    });

    const selectBtn = document.getElementById("select-central-node-btn");
    selectBtn.onclick = () => {
        console.log("Attempting to select, current:", selectedCentralNode);
        if (!selectedCentralNode) {
            alert("Please select a central node first!");
            return;
        }
        centralNode = selectedCentralNode;
        console.log("Final central node:", centralNode);
        popup.style.display = "none";
        insertObjectSummaryTables();
    };

    document.getElementById("close-central-node").onclick = () => {
        popup.style.display = "none";
    };

    popup.style.display = "block";
}

function closeCentralNodePicker() {
    document.querySelectorAll("#object-summary-selection input[type=checkbox]:checked").map(cb => cb.value);
}

function extractSentenceContaining(text, value) {
    const sentences = text.split(/(?<=[.!?])/);
    return sentences.find(s => s.includes(value))?.trim() || value;
}

async function sendGraphToBackend() {
    if(!centralNode) {
        console.log("No central node selected.");
        return;
    }
    
    const objects = [];
    
    Object.keys(taggedObjects).forEach(type => {
        taggedObjects[type].forEach(obj => {
            if (obj.value !== centralNode.value) {
                objects.push({
                    name: obj.value,
                    type: type.replace(/\s/g, "")
                });
            }
        });
    });
    
    const payload = {
        centralNode: centralNode.value,
        objects: objects,
    };
    
    console.log("GRAPH PAYLOAD:", payload);
    
    try {
        const response = await fetch("https://localhost:7225/knowledge-graph", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
            },
            body: JSON.stringify(payload)
        });
        
        const result = await response.text();
        console.log("GRAPH RESULT:", result);
        
        //opet browser 
        window.open(
            "http://localhost:7474/browser/?cmd=MATCH%20(n)-[r]->(m)%20RETURN%20n,r,m",
            "_blank"
        );
        
    } catch (error) {
        console.log("Graph API Error:", error);
    }
}

function getSelectedObjectTypes() {
    const checkboxes = document.querySelectorAll(
        "#object-summary-selection input[type=checkbox]:checked"
    );

    return Array.from(checkboxes).map(cb => cb.value);
}