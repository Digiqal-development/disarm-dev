const url = "https://disarm.housepilot.de"


export async function getClauses(sentence) {
    try {
        const result = await $.ajax({
        type: 'POST',
        dataType: 'json',
        url: url + "/clauses",
        contentType: 'application/json',
        data: JSON.stringify({
          sentence: sentence.text,
          result: 'string'
        })
      });
      return result;
    } catch (error) {
      console.error('Error fetching clauses:', error);
      throw error; 
    }
  }

export async function getTechniques() {
    try {
        const result = await $.ajax({
            type: 'GET',
            headers: {
              'Cache-Control': 'no-cache',
              'Pragma': 'no-cache',
              'Expires': '0'
            },
            dataType: "json",
            url: url + "/techniques"
        });
        return result;
    } catch (error) {
        console.error('Error fetching techniques:', error);
    }
}

export async function searchTechniques() {
    try {
        const result = await $.ajax({
            type: 'GET',
            headers: {
              'Cache-Control': 'no-cache',
              'Pragma': 'no-cache',
              'Expires': '0'
            },
            dataType: "json",
            url: url + "/tags"
        });
        return result;
    } catch (error) {
        console.error('Error fetching techniques:', error);
    }
}
