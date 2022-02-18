var id = '1TFdXJ3VVi-c6pn36PMltkMJDympYS3TDw586uzc8Abw'
function Caixa(){
  var spreadsheet = SpreadsheetApp.getActive();
  var ativador = spreadsheet.getSheetByName('Base de preços')
  var linha = 2
  while(ativador.getRange(linha,17).isBlank() == false){
    linha++
  }
  var fornec = ativador.getRange(2,17,linha-2).getValues();
  var Form = HtmlService.createTemplateFromFile("SaídaFCaixa");
  Logger.log(fornec);
  Form.fornec = fornec.map(function(r){return r[0];});
  var Mostrar = Form.evaluate()
  Mostrar.setTitle("Caixa").setHeight(500).setWidth(600)
  SpreadsheetApp.getUi().showModalDialog(Mostrar,"Caixa")
}

function EntradaFCaixa(){
  var Form = HtmlService.createTemplateFromFile("EntradaFCaixa")
  var Mostrar = Form.evaluate()
  Mostrar.setTitle("Caixa").setHeight(500).setWidth(400)
  SpreadsheetApp.getUi().showModalDialog(Mostrar,"Caixa")
}

function Produtos(){
  var Form = HtmlService.createTemplateFromFile("Produtos")
  var Mostrar = Form.evaluate()
  Mostrar.setTitle("Cadastro de Produtos").setHeight(500).setWidth(400)
  SpreadsheetApp.getUi().showModalDialog(Mostrar,"Cadastro de Produtos")
}

function Sangriaatv(){
  var spreadsheet = SpreadsheetApp.getActive();
  var ativador = spreadsheet.getSheetByName('Base de preços')
  var linha = 2
  while(ativador.getRange(linha,15).isBlank() == false){
    linha++
  }
  var guias = ativador.getRange(2,15,linha-2).getValues();
  var Form = HtmlService.createTemplateFromFile("Sangria");
  Logger.log(guias);
  Form.guias = guias.map(function(r){return r[0];});
  var Mostrar = Form.evaluate();
  Mostrar.setTitle("Sangria").setHeight(330).setWidth(250);
  SpreadsheetApp.getUi().showModalDialog(Mostrar,"Sangria");
}

function Pesquisar(){
  var Form = HtmlService.createTemplateFromFile("Pesquisar1")
  var Mostrar = Form.evaluate()
  Mostrar.setTitle("Pesquisa").setHeight(500).setWidth(600)
  SpreadsheetApp.getUi().showModalDialog(Mostrar,"Pesquisa")
}

function CaixaFunc(){
  var spreadsheet = SpreadsheetApp.getActive();
  var produtos = spreadsheet.getSheetByName('Base de Preços')
  var caixaat = spreadsheet.getSheetByName('Tela do Caixa')
  var linha = 2
  while(produtos.getRange(linha,6).isBlank() == false){
    linha++
  }
  var bebidas = produtos.getRange(2,6,linha-2,1).getValues()
  linha = 2
  while(produtos.getRange(linha,7).isBlank() == false){
    linha++
  }
  var refeicao = produtos.getRange(2,7,linha-2,1).getValues()
  linha = 2
  while(produtos.getRange(linha,8).isBlank() == false){
    linha++
  }
  var outros = produtos.getRange(2,8,linha-2,1).getValues()
  var Form = HtmlService.createTemplateFromFile("CaixaFuncionarios")
  Form.bebidas = bebidas.map(function(r){return r[0];});
  Form.refeicao = refeicao.map(function(r){return r[0];});
  Form.outros = outros.map(function(r){return r[0];});
  var Mostrar = Form.evaluate()
  Mostrar.setTitle("Registro de Comanda").setHeight(500).setWidth(800)
  SpreadsheetApp.getUi().showModalDialog(Mostrar,"Registro de Comanda")
  }

function Vendas(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var produtos = spreadsheet.getSheetByName('Base de Preços')
  var linha = 1
  while(produtos.getRange(linha,2).isBlank() == false){
    linha = linha+1
  }
  var list = produtos.getRange(2,2,linha-2,1).getValues()
  var Form = HtmlService.createTemplateFromFile("Vendas")
  Logger.log(list)
  Form.list = list.map(function(r){return r[0];});
  var Mostrar = Form.evaluate()
  Mostrar.setTitle("Cadastro de Venda").setHeight(500).setWidth(400)
  SpreadsheetApp.getUi().showModalDialog(Mostrar,"Cadastro de Venda")
}

function enviar(Dados) {
  var spreadsheet = SpreadsheetApp.openById(id)
  var sheet = spreadsheet.getSheetByName('Fluxo de Caixa')
  sheet.getRange('L1').activate()
  sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
  sheet.getActiveCell().offset(1,0).activate()
  var linha = sheet.getActiveCell().getRow()
  var data = new Date();
  var dia = String(data.getDate()). padStart(2,'0')
  var mes = String(data.getMonth() + 1). padStart(2,'0')
  var ano = data. getFullYear()
  var data = dia + '/' + mes + '/' + ano
  sheet.getRange(linha,11).setValue(data)
  sheet.getRange(linha, 12).setValue([Dados.ent])
  sheet.getRange(linha, 13).setValue([Dados.forn])
  sheet.getRange(linha, 14).setValue([Dados.desc])
  sheet.getRange(linha, 15).setValue([Dados.val]) 
  while(sheet.getRange(linha, 12).isBlank() == false) {
    linha = linha + 1
  } 
  sheet.getRange(2,11,linha-2,5).activate();
  sheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}
function onEditAction(e) {
  var ss = e.source.getActiveSheet()
  var range = e.range.getA1Notation()
  if(ss.getName() != 'Fluxo de Caixa' || range != 'G22') {return}
  Caixa()
  ss.getRange('G22').clearContent()
  }
function enviarProduto(Dados) {
  var spreadsheet = SpreadsheetApp.openById(id)
  var sheet = spreadsheet.getSheetByName('Base de Preços')
  sheet.getRange('B1').activate()
  sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
  sheet.getActiveCell().offset(1,0).activate()
  var linha = sheet.getActiveCell().getRow()
  sheet.getRange(linha, 1).setValue(linha-1)
  sheet.getRange(linha, 2).setValue([Dados.nome])
  sheet.getRange(linha, 3).setValue([Dados.val])
  sheet.getRange(linha, 4).setValue([Dados.grupo])
  if(Dados.grupo == "Bebidas"){
    sheet.getRange('F2').activate()
    sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
    sheet.getActiveCell().offset(1,0).activate()
    var linha = sheet.getActiveCell().getRow()
    sheet.getRange(linha, 6).setValue([Dados.nome])
  }
   if(Dados.grupo == "Restaurante"){
    sheet.getRange('G2').activate()
    sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
    sheet.getActiveCell().offset(1,0).activate()
    var linha = sheet.getActiveCell().getRow()
    sheet.getRange(linha, 7).setValue([Dados.nome])
  }
  if(Dados.grupo == "Diversos"){
    sheet.getRange('H2').activate()
    sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
    sheet.getActiveCell().offset(1,0).activate()
    var linha = sheet.getActiveCell().getRow()
    sheet.getRange(linha, 8).setValue([Dados.nome])
  }
  sheet.getRange('F:H').setHorizontalAlignment("center")
  
  }

function enviarVenda(Dados){
  var spreadsheet = SpreadsheetApp.openById(id)
  var sheet = spreadsheet.getSheetByName('Vendas')
  var data = new Date();
  var dia = String(data.getDate()). padStart(2,'0')
  var mes = String(data.getMonth() + 1). padStart(2,'0')
  var ano = data. getFullYear()
  var data = dia + '/' + mes + '/' + ano
  var hora = data.getHours
  if(hora<=15){
    var turno = 'Diurno'
  } else{
    var turno = 'Noturno'
  }
  Logger.log(turno)
  var produtos = spreadsheet.getSheetByName('Base de Preços')
  lin = 1
  while(produtos.getRange(lin,2).isBlank() == false){
    lin = lin+1
  }
  // Pegando uma lista com os nomes dos produtos e seus valores
  var list = produtos.getRange(2,2,lin-2,1).getValues()
  var valor = produtos.getRange(2,3,lin-2,1).getValues()
  for(var i=0;i<list.length;i++){
    for(var j=0;j<list[i].length;j++){
      if(list[i][j]==Dados.prod){
        var val = valor[i][j]
      }
    }
  }
  // Salvando dados nas células
  sheet.getRange('A1').activate()
  sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
  sheet.getActiveCell().offset(1,0).activate()
  var linha = sheet.getActiveCell().getRow()
  sheet.getRange(linha, 1).setValue(data)
  sheet.getRange(linha, 2).setValue(turno)
  sheet.getRange(linha, 3).setValue([Dados.prod])
  sheet.getRange(linha, 4).setValue([Dados.qnt])
  sheet.getRange(linha, 5).setValue(val*[Dados.qnt])
}

function pesquisa(DadosP) {
  var data = DadosP.data
  var splitdata = data.split('-')
  var refdata = (splitdata[2]+'/'+splitdata[1]+'/'+splitdata[0])
  var sheet = SpreadsheetApp.getActiveSheet();
  var linha = 2
  while(sheet.getRange(linha,11).getValue() != refdata){
    linha++
  }
  var seta = linha
  while(sheet.getRange(linha,11).getValue() == refdata){
    linha++
  }
  sheet.getRange(seta,11,linha-seta,5).activate().setBackground('red')
};

function removerfiltro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('K1:O1').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveSheet().getFilter().remove();
  spreadsheet.getRange('K:K').setNumberFormat('dd/MM/yyyy')
};
function Aporte(){
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Tela do Caixa'), true);
  var Form = HtmlService.createTemplateFromFile("Aporte")
  var Mostrar = Form.evaluate()
  Mostrar.setTitle("Aporte").setHeight(300).setWidth(250)
  SpreadsheetApp.getUi().showModalDialog(Mostrar,"Aporte")
}
function entradacomanda(products,quanti,formapg,tipovenda){
  var spreadsheet = SpreadsheetApp.openById(id)
  var sheet = spreadsheet.getSheetByName('Tela do Caixa')
  var produtos = spreadsheet.getSheetByName('Base de Preços')
  var lin = produtos.getLastRow()
  var valores = produtos.getRange(2,2,lin-1,3).getValues()
  sheet.getRange('B9').activate()
  sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
  sheet.getActiveCell().offset(1,0).activate()
  var linha = sheet.getActiveCell().getRow()
  var seta = linha
  sheet.getRange(linha, 1).setValue('C' + seta)
  sheet.getRange(linha, 5).setValue(formapg)
  sheet.getRange(linha, 6).setValue(tipovenda)
  var comandav = 0
  for(i=0; i<products.length; i++ ){
    sheet.getRange(linha, 2).setValue(products[i])
    sheet.getRange(linha, 3).setValue(quanti[i])
    for(j=0;j<valores.length;j++){
      if(products[i]==valores[j][0]){
        sheet.getRange(linha, 4).setValue(valores[j][1]*quanti[i])
        var comandav = comandav + valores[j][1]*quanti[i]
      }
    }
    linha = linha+1
    }
  if (formapg == "Dinheiro"){ 
      var ant = sheet.getRange('K17').getValue() 
      var novo = comandav
      sheet.getRange('K17').setValue(novo + ant)}
  if (formapg == "Crédito"){ 
      var ant = sheet.getRange('K18').getValue() 
      var novo = comandav
      sheet.getRange('K18').setValue(novo + ant)}
  if (formapg == "Débito"){ 
      var ant = sheet.getRange('K19').getValue() 
      var novo = comandav
      sheet.getRange('K19').setValue(novo + ant)}
  if (formapg == "Pix"){ 
      var ant = sheet.getRange('K20').getValue() 
      var novo = comandav
      sheet.getRange('K20').setValue(novo + ant)
    }

  sheet.getRange(seta,1,products.length,1).activate().mergeVertically();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(seta, 5,products.length, 2).activate().mergeVertically();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center').setVerticalAlignment('middle');
  var linha = sheet.getRange('B9').getRow()
  while(sheet.getRange(linha, 2).isBlank() == false) {
    linha = linha + 1
  } 
  sheet.getRange(11,1,linha-11,6).activate();
  sheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}
function Save(Dados){
  var spreadsheet = SpreadsheetApp.openById(id)
  var sheet = spreadsheet.getSheetByName('Tela do Caixa')
  sheet.getRange('B9').activate()
  sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
  sheet.getActiveCell().offset(1,0).activate()
  var linha = sheet.getActiveCell().getRow()
  sheet.getRange(linha, 1).setValue('AP' + linha).setHorizontalAlignment('center')
  sheet.getRange(linha, 4).setValue([Dados.Quantidade])
  sheet.getRange(linha, 2).setValue([Dados.Desc])
  var ant = sheet.getRange('K15').getValue() 
  var novo = sheet.getRange(linha, 4).getValue()
  sheet.getRange('K15').setValue(novo + ant)
  while(sheet.getRange(linha, 2).isBlank() == false) {
    linha = linha + 1
  } 
  sheet.getRange(11,1,linha-11,6).activate();
  sheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}

function Sangria(Dados){
  var spreadsheet = SpreadsheetApp.openById(id)
  var sheet = spreadsheet.getSheetByName('Tela do Caixa')
  sheet.getRange('B9').activate()
  sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
  sheet.getActiveCell().offset(1,0).activate()
  var linha = sheet.getActiveCell().getRow()
  sheet.getRange(linha, 1).setValue('SG' + linha).setHorizontalAlignment('center')
  sheet.getRange(linha, 4).setValue([Dados.Quantidade])
  sheet.getRange(linha, 2).setValue([Dados.Desc])
  sheet.getRange(linha, 5).setValue([Dados.Nome]).setHorizontalAlignment('center')
  var ant = sheet.getRange('K16').getValue() 
  var novo = sheet.getRange(linha, 4).getValue()
  sheet.getRange('K16').setValue(novo + ant)
   while(sheet.getRange(linha, 2).isBlank() == false) {
    linha = linha + 1
  } 
  sheet.getRange(11,1,linha-11,5).activate();
  sheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}

function Fechat(){
  var spreadsheet = SpreadsheetApp.openById(id)
  var sheet = spreadsheet.getSheetByName('Tela do Caixa')
  var linha = 11
  var seta = linha
  while(sheet.getRange(linha, 2).isBlank() == false) {
    linha = linha + 1}
  var outro = spreadsheet.getSheetByName('Dados')
  var ahnil = 2
  while(outro.getRange(ahnil, 13).isBlank() == false) {
    ahnil = ahnil + 1}
  var point = outro.getRange(ahnil,13)
  var dia = sheet.getRange('C4').getValue()
  var turno = sheet.getRange('C6').getValue()
  outro.getRange(ahnil, 11).setValue(turno)
  outro.getRange(ahnil, 10).setValue(dia)
  sheet.getRange(seta, 2, linha-seta, 5).moveTo(point)
  outro.getRange('P:Q').breakApart()
  sheet.getRange('A11:A').deleteCells(SpreadsheetApp.Dimension.ROWS)
  for (i=ahnil+1; i<linha-seta+ahnil; i++){
    outro.getRange(i, 11).setValue(turno)
    outro.getRange(i, 10).setValue(dia)
    if (outro.getRange(i,17).isBlank() == true){
        outro.getRange(i,17).setValue(outro.getRange(i-1,17).getValue())}
    if (outro.getRange(i,16).isBlank() == true){
        outro.getRange(i,16).setValue(outro.getRange(i-1,16).getValue())}
    Logger.log(i)
  }}
  

