const url = "https://disarm2.housepilot.de";
const url1 = "https://localhost:7225";

export async function getClauses(sentence) {
  try {
    const result = await $.ajax({
      type: "POST",
      dataType: "json",
      url: url + "/clauses",
      contentType: "application/json",
      data: JSON.stringify({
        sentence: sentence.text,
        result: "string",
      }),
    });
    return result;
  } catch (error) {
    console.error("Error fetching clauses:", error);
    throw error;
  }
}

export async function getTechniques() {
  try {
    const result = await $.ajax({
      type: "GET",
      headers: {
        "Cache-Control": "no-cache",
        Pragma: "no-cache",
        Expires: "0",
      },
      dataType: "json",
      url: url + "/techniques",
    });
    return result;
  } catch (error) {
    console.error("Error fetching techniques:", error);
  }
}

export async function searchTechniques() {
  try {
    const result = await $.ajax({
      type: "GET",
      headers: {
        "Cache-Control": "no-cache",
        Pragma: "no-cache",
        Expires: "0",
      },
      dataType: "json",
      url: url + "/tags",
    });
    return result;
  } catch (error) {
    console.error("Error fetching techniques:", error);
  }
}

export function createUrlWithValuesArray(endpoint, inputArray) {
    const jsonArrayString = JSON.stringify(inputArray);
    const encodedJson = encodeURIComponent(jsonArrayString);
    const queryParameter = `value=${encodedJson}`;
    return `${url}${endpoint}?${queryParameter}`;
}
