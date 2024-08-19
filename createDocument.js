function createNewGoogleDocs() {
  // Get the current sheet and its data
  var spreadsheet = SpreadsheetApp.getActive();
  var sheetDoc = spreadsheet.getSheetByName("Preencher Documento");
  var start_sheet = spreadsheet.getSheetByName('Start')
  var initial_counting_number = start_sheet.getRange('B5').getValue().toString();
  var row_number = sheetDoc.getLastRow().toString();
  var row_number = row_number.replace(".0","");
  var file_name = start_sheet.getRange('B6').getValue().toString();
  Logger.log(file_name);
  Logger.log(initial_counting_number);  
// Get your template and destination folder
  var template = DriveApp.getFileById("1UAqxZH-Ltd23_W5KX-5Ya-8vdx7YCeZ2TZC4lU-HS8w");
  var destinationFolder = DriveApp.getFolderById('1dFc8kKmtVsoUvgUhM9YuGkqsy6_JW87A');
  const copy = template.makeCopy(file_name, destinationFolder);
  var doc = DocumentApp.openById(copy.getId());
  var body = doc.getBody(); 
  // Loop through each row in the sheet
  for(i=2; i<=row_number;i++){
    var number = sheetDoc.getRange('A'+i).getValue().toString();
    var description = sheetDoc.getRange('D'+i).getValue();
    var characteristics = sheetDoc.getRange('E'+i).getValue();
    var equipaments = sheetDoc.getRange('F'+i).getValue();
    var quantification =sheetDoc.getRange('G'+i).getValue();
    var measurement = sheetDoc.getRange('H'+i).getValue();
    var execution = sheetDoc.getRange('I'+i).getValue();
    var complements =sheetDoc.getRange('J'+i).getValue();

    equipaments = dealwithvoid(equipaments);
    quantification = dealwithvoid(quantification);
    measurement = dealwithvoid(measurement);
    execution = dealwithvoid(execution);
    complements = dealwithvoid(complements);

    const item = new Item(number, description, characteristics, equipaments, quantification, measurement, execution, complements);
    item.createtitle(initial_counting_number, body);

    if(characteristics != ""){
      item.addParagraph(body);
    }
    else{
      continue;
    }
      
  }
  doc.saveAndClose();
  var url = doc.getUrl();
  start_sheet.getRange('E9').setValue(url);
}

class Item{
  constructor(number, description, characteristics, equipaments,quantification, measurement, execution, complements){
    this.number = number;
    this.description = description;
    this.characteristics = characteristics;
    this.equipaments = equipaments;
    this.quantification = quantification;
    this.measurement = measurement;
    this.execution = execution;
    this.complements =complements;
    this.pending = "Não se aplica."
  }
  createtitle(initial_counting_number, body){
    var title = initial_counting_number + '.' + this.number+ ' ' + this.description;
    title = title.toUpperCase();
    var header = body.appendParagraph(title);
    header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('');
  }
  addParagraph(body){
    // Append a styled paragraph.
    var par = body.appendParagraph('A. ITENS E SUAS CARACTERÍSTICAS');
    par.setBold(true);
    var text = body.appendParagraph(this.characteristics);
    text.setBold(false);
    body.appendParagraph('');
    var par = body.appendParagraph('B. EQUIPAMENTOS');
    par.setBold(true);
    text = body.appendParagraph(this.equipaments);
    text.setBold(false);
    body.appendParagraph('');
    var par = body.appendParagraph('C. CRITÉRIOS PARA QUANTIFICAÇÃO DOS SERVIÇOS');
    par.setBold(true);
    text = body.appendParagraph(this.quantification);
    text.setBold(false);
    body.appendParagraph('');
    var par = body.appendParagraph('D. CRITÉRIOS DE AFERIÇÃO');
    par.setBold(true);
    text = body.appendParagraph(this.measurement);
    text.setBold(false);
    body.appendParagraph('');
    var par = body.appendParagraph('E. EXECUÇÃO');
    par.setBold(true);
    text = body.appendParagraph(this.measurement);
    text.setBold(false);
    body.appendParagraph('');
    var par = body.appendParagraph('F. INFORMAÇÕES COMPLEMENTARES');
    par.setBold(true);
    text = body.appendParagraph(this.complements);
    text.setBold(false);
    body.appendParagraph('');
    var par = body.appendParagraph('G. PENDÊNCIAS');
    par.setBold(true);
    text= body.appendParagraph(this.pending);
    text.setBold(false);
    body.appendParagraph('');
        
  }

}
function dealwithvoid(string){
  if(string==""){
    string="Não se aplica.";
  }
  return string;
}
