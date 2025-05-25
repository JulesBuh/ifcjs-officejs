//$("#get-paragraph").click(() => tryCatch(getParagraph));
$("#addDocument").click(() => tryCatch(getParagraph));

$(document).ready(function() {
  $('[data-bs-toggle="tooltip"]').tooltip();
});

setAppContext("docx");

/** sets that application context by passing it the open document's file extension*/
function setAppContext(format = "docx") {
  const appcontext = "-" + getContext(format);

  function getContext(format) {
    switch (format) {
      case "":
        return "goldlink";
      case "docx":
        return "word";
      case "docm":
        return "word";
      case "doc":
        return "word";
      case "dot":
        return "word";
      case "dotm":
        return "word";
      case "dotx":
        return "word";
      case "odt":
        return "word";
      case "rtf":
        return "word";
      case "pdf":
        return "pdf";
      case "xlsx":
        return "excel";
      case "xltx":
        return "excel";
      case "xltm":
        return "excel";
      case "xlmx":
        return "excel";
      case "xls":
        return "excel";
      case "xlsb":
        return "excel";
      case "xlt":
        return "excel";
      case "xlm":
        return "excel";
      case "ods":
        return "excel";
      case "csv":
        return "excel";
      case "tsv":
        return "excel";
      case "xml":
        return "excel";
      case "ppt":
        return "powerpoint";
      case "pps":
        return "powerpoint";
      case "pptx":
        return "powerpoint";
      case "ppsx":
        return "powerpoint";
      case "pptm":
        return "powerpoint";
      case "odp":
        return "powerpoint";
      case "one":
        return "onenote";
      default:
        return "goldlink";
    }
  }

  //const contextcolor = document.styleSheets[0].cssRules[0].styleMap.get("--gl" + appcontext);
  //const contexticon = "fa-file" + appcontext;
  //document.styleSheets[0].cssRules[0].styleMap.set("--gl-applicationcontext", contextcolor);
  //document.styleSheets[0].cssRules[0].styleMap.set("--gl-applicationcontext-transp", contextcolor + "50");
  //$(".appcontext").addClass(contexticon);
  $(".appcontext").addClass("fa-file"); //adds the default fa icon in addition to the context icon
}

async function loadFileName() {
  return new Promise((resolve) => {
    Office.context.document.getFilePropertiesAsync(null, (res) => {
      if (res && res.value && res.value.url) {
        let name = res.value.url.substr(res.value.url.lastIndexOf("\\") + 1);
        resolve({ directory: res.value.url, filename: name });
      }
      resolve("none");
    });
  });
}

async function getParagraph() {
  await Word.run(async (context) => {
    // The collection of paragraphs of the current selection returns the full paragraphs contained in it.
    const paragraph = context.document.getSelection();
    paragraph.load();
    const props = context.document.properties;
    const doc = context.document;
    props.load();
    doc.load();

    await context.sync();
    let moreprops = await loadFileName();
    // write this to the ifc model
    addNewObject(0, document.getElementById("expressIDLabel").value, paragraph.text, props, moreprops);
    //console.log(props, doc, await loadFileName());
    //console.log(paragraph.text);
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

/*---------------------------------------------------*/
class AccentItem {
  constructor(id, name, color) {
    this.id = id;
    this.name = name;
    this.color = color;
  }
}
var AccentColours = { Accent1: null, Accent2: null, Accent3: null, Accent4: null, Accent5: null, Accent6: null };
async function registerAccentStyles() {
  AccentColours["Accent1"] = Office.context.officeTheme.bodyBackgroundColor;
  AccentColours["Accent2"] = Office.context.officeTheme.bodyForegroundColor;
  AccentColours["Accent3"] = Office.context.officeTheme.controlBackgroundColor;
  AccentColours["Accent4"] = Office.context.officeTheme.controlForegroundColor;

  //console.log(AccentColours);
}

registerAccentStyles();
$(".body").css("background-color", AccentColours["Accent1"]);
/*----------------------------------------------------*/
/** ifcObjectBuilder */
function addNewObject(modelID = 0, object = document.getElementById("expressIDLabel").value, text = "", props, doc) {
  const nextexpressId = getLastExpressId();
  let currentDate = new Date();
  let applicationName = "BIMDoc" + document.URL;
  applicationName = applicationName.replace("'", "");
  let docname = (doc.filename = "" ? "######" : doc.filename + "#" + nextexpressId);
  docname = docname.replace("'", "");
  let doctitle = props.title == "" ? "Untitled" : props.title;
  doctitle = doctitle.replace("'", "");
  let doclocation = doc.directory == "" ? "C:\\" : doc.directory;
  doclocation = doclocation.replace("'", "");
  let company = props.company == "" ? "ORG" : props.company;
  company = company.replace("'", "");
  let author = props.lastAuthor == "" ? "AUTH" : props.lastAuthor;
  author = author.replace("'", "");
  let firstAuthor = props.author == "" ? "AUTH" : props.author;
  firstAuthor = firstAuthor.replace("'", "");
  let manager = props.manager == "" ? "AUTH" : props.manager;
  manager = manager.replace("'", "");
  let subject = props.subject == "" ? "" : props.subject;
  subject = subject.replace("'", "");
  let keywords = props.keywords == "" ? "" : props.keywords;
  keywords = keywords.replace("'", "");
  let comments = props.comments == "" ? "" : props.comments;
  comments = comments.replace("'", "");
  let format = props.format == "" ? "docx" : props.format;
  format = format.replace("'", "");
  let revision = props.revisionNumber == "" ? "" : props.revisionNumber;
  revision = revision.replace("'", "");
  let savedate = props.lastSaveTime == "" ? currentDate.toISOString() : props.lastSaveTime;
  //savedate = savedate.replace("'", "");

  let org = new IfcOrganization(
    nextexpressId,
    IFCORGANIZATION,
    { type: 1, value: "mybusiness.microsoft.com" },
    { type: 1, value: "mybusiness" },
    { type: 1, value: "BIM development tools" },
    { type: 1, value: "BIM developer" },
    { type: 1, value: "julesbuh@mybusiness.microsoft.com" }
  );

  ifc.state.api.WriteLine((modelID = 0), org);

  // Build Application Details
  let Application = new IfcApplication(
    nextexpressId + 1,
    IFCAPPLICATION,
    org,
    { type: 1, value: "0.0" },
    { type: 1, value: applicationName },
    { type: 1, value: AddDashes(md5("BIMDoc" + "0.0"), 3) }
  );

  ifc.state.api.WriteLine((modelID = 0), Application);

  //Build and person organisation
  let PersonOrg = new IfcPersonAndOrganization(
    nextexpressId + 2,
    IFCPERSONANDORGANIZATION,
    { type: 1, value: "The End User" },
    { type: 1, value: company == "" ? "The End Users Organization" : company },
    { type: 1, value: "The End Users Role" }
  );

  ifc.state.api.WriteLine((modelID = 0), PersonOrg);

  //Build an owner
  let Owner = new IfcOwnerHistory(
    nextexpressId + 3,
    IFCOWNERHISTORY,
    PersonOrg,
    Application,
    { type: 1, value: "READONLY" },
    { type: 1, value: "ADDED" },
    { type: 1, value: "" + currentDate.toISOString() },
    PersonOrg,
    Application,
    { type: 1, value: "" + currentDate.toISOString() }
  );

  ifc.state.api.WriteLine((modelID = 0), Owner);

  // Builds the document
  var Doc = new IfcDocumentInformation(
    nextexpressId + 4,
    IFCDOCUMENTINFORMATION,
    { type: 1, value: docname + "" },
    { type: 1, value: doctitle + "" },
    { type: 1, value: text },
    { type: 1, value: doclocation },
    { type: 1, value: subject },
    { type: 1, value: keywords },
    { type: 1, value: comments },
    { type: 1, value: "" + revision },
    { type: 1, value: company == "" ? "The End Users Organization" : company + ", " + firstAuthor },
    { type: 1, value: author == "" ? firstAuthor : firstAuthor + ", " + author + ", " + manager },
    { type: 1, value: "" + currentDate.toISOString() },
    { type: 1, value: currentDate.toISOString() },
    { type: 1, value: format },
    { type: 1, value: "" + currentDate.toISOString() },
    { type: 1, value: "" + currentDate.toISOString() },
    { type: 1, value: "CONFIDENTIAL" },
    { type: 1, value: "WIP" }
  );
  //console.log(Doc);

  ifc.state.api.WriteLine((modelID = 0), Doc);

  // Builds the relationship to the object
  var ObjRelDoc = new IfcRelAssociatesDocument(
    nextexpressId + 5,
    IFCRELASSOCIATESDOCUMENT,
    { type: 1, value: "" + GuidToIfc(md5(Doc.Identification)) },
    Owner,
    { type: 1, value: "" + "Document Section - Page Number" },
    { type: 1, value: "" + text },
    [{ type: 5, value: Number(object) }],
    Doc
  );

  ifc.state.api.WriteLine((modelID = 0), ObjRelDoc);
  //reload properties panel
  getPropertyWithExpressId(modelID, Doc.expressID);
  //set the selector to new object's id
  document.getElementById("expressIDLabel").value = Doc.expressID;

  //getPropertyWithExpressId(modelID, Doc.expressID);
}
