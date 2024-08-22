
//test
async function test12(){
    await Word.run(async (context) => {
      const originalXml =
        "<Locations><Location>Juan</Location><Location>Hong</Location><Location>Sally</Location></Locations>";
      const customXmlPart = context.document.customXmlParts.add(originalXml);
      customXmlPart.load("id");
      const xmlBlob = customXmlPart.getXml();
  
      await context.sync();
  
      const readableXml = addLineBreaksToXML(xmlBlob.value);
      console.log("Added custom XML part:");
      console.log(readableXml);
  
      // Store the XML part's ID in a setting so the ID is available to other functions.
      const settings = context.document.settings;
      settings.add("ContosoReviewXmlPartId", customXmlPart.id);
  
      await context.sync();
    });
  }
  
  async function test1(){
    await Word.run(async (context) => {
    var doc1 = context.document.getSelection().getTextRanges(['\n', '.', '?'], false);
        
    context.load(doc1);
    doc =  await context.sync().then(() => { return doc1.items[0]})
  
    //doc.text = "nesto drugo"
    
    //send sentence to backend for clauses extraction
    //await getClauses(doc).then(()=>  console.log(clausesGlobal));
  
    var clauses = await getClauses(doc);
    var string = clauses.result
  
  
    let jsonString = string.replace(/'/g, '"'); // Removes leading and trailing single quotes
  
  
  // Step 2: Use JSON.parse() to convert the JSON string to an array
    let array = JSON.parse(jsonString);
  
  
    
    var rangeToHighlight = doc.search(array[0]);
  
    rangeToHighlight.load('items');
    await context.sync();
    console.log(rangeToHighlight.items[0].text);
    rangeToHighlight.items[0].font.color = 'red';
  
    //var clausesArray = JSON.parse(clauses.result)
    //console.log(clausesArray[0])
    await context.sync();
    })
  }
  
  
  function addLineBreaksToXML( xmlBlob) {
    const replaceValue = new RegExp(">");
    return xmlBlob.replace(/></g, "> <");
  }
  
  async function test2(){
    // Queries a custom XML part for elements matching the search terms.
    await Word.run(async (context) => {
      const settings = context.document.settings;
      const xmlPartIDSetting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");
  
      await context.sync();
  
      if (xmlPartIDSetting.value) {
        const customXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);
        const xpathToQueryFor = "/Locations/Location";
        const clientResult = customXmlPart.query(xpathToQueryFor, {
          contoso: "http://schemas.contoso.com/review/1.0"
        });
  
        await context.sync();
  
        console.log(`Queried custom XML part for ${xpathToQueryFor} and found ${clientResult.value.length} matches:`);
        for (let i = 0; i < clientResult.value.length; i++) {
          console.log(clientResult.value[i]);
        }
      } else {
        console.warn("Didn't find custom XML part to query");
      }
    });
  }
  
  
  
  
  
  
  