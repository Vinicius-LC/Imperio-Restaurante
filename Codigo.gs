var id = '1t8AncLhqZ_hIjVNdpc7kWLh3tR-nMvMZbVXXZxB_L4o'
function Caixa() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ativador = spreadsheet.getSheetByName('Base de preços')
  var linha = 2
  while (ativador.getRange(linha, 17).isBlank() == false) {
    linha++
  }
  var fornec = ativador.getRange(2, 17, linha - 2).getValues();
  var setores = ativador.getRange(2,19,(ativador.getRange(2,19).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()-1),1).getValues()
  var Form = HtmlService.createTemplateFromFile("SaídaFCaixa");
  Logger.log(fornec);
  Form.setores = setores.map(function (r) { return r[0]; });
  Form.fornec = fornec.map(function (r) { return r[0]; });
  var Mostrar = Form.evaluate()
  Mostrar.setTitle("Caixa").setHeight(500).setWidth(600)
  SpreadsheetApp.getUi().showModalDialog(Mostrar, "Caixa")
}
function Fecha() {
  var spreadsheet = SpreadsheetApp.getActive()
  var Form = HtmlService.createTemplateFromFile("Fechamento");
  var Mostrar = Form.evaluate()
  Mostrar.setTitle("Fechameto de Caixa").setHeight(100).setWidth(300)
  SpreadsheetApp.getUi().showModalDialog(Mostrar, "Fechameto de Caixa")
}
function Manutencaoatv() {
  var Form = HtmlService.createTemplateFromFile("Manutenção");
  var Mostrar = Form.evaluate()
  Mostrar.setTitle("Extras").setHeight(500).setWidth(600)
  SpreadsheetApp.getUi().showModalDialog(Mostrar, "Extras")
}

function EntradaFCaixa() {
  var Form = HtmlService.createTemplateFromFile("EntradaFCaixa")
  var Mostrar = Form.evaluate()
  Mostrar.setTitle("Caixa").setHeight(500).setWidth(400)
  SpreadsheetApp.getUi().showModalDialog(Mostrar, "Caixa")
}

function Produtos() {
  var Form = HtmlService.createTemplateFromFile("Produtos")
  var Mostrar = Form.evaluate()
  Mostrar.setTitle("Cadastro de Produtos").setHeight(500).setWidth(400)
  SpreadsheetApp.getUi().showModalDialog(Mostrar, "Cadastro de Produtos")
}

function Sangriaatv() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ativador = spreadsheet.getSheetByName('Base de preços')
  var linha = 2
  while (ativador.getRange(linha, 15).isBlank() == false) {
    linha++
  }
  var guias = ativador.getRange(2, 15, linha - 2).getValues();
  var Form = HtmlService.createTemplateFromFile("Sangria");
  Logger.log(guias);
  Form.guias = guias.map(function (r) { return r[0]; });
  var Mostrar = Form.evaluate();
  Mostrar.setTitle("Sangria").setHeight(330).setWidth(250);
  SpreadsheetApp.getUi().showModalDialog(Mostrar, "Sangria");
}

function Pesquisar() {
  var Form = HtmlService.createTemplateFromFile("Pesquisar1")
  var Mostrar = Form.evaluate()
  Mostrar.setTitle("Pesquisa").setHeight(500).setWidth(600)
  SpreadsheetApp.getUi().showModalDialog(Mostrar, "Pesquisa")
}

function CaixaFunc() {
  var spreadsheet = SpreadsheetApp.getActive();
  var produtos = spreadsheet.getSheetByName('Base de Preços')
  var caixaat = spreadsheet.getSheetByName('Tela do Caixa')
  var linha = 2
  while (produtos.getRange(linha, 6).isBlank() == false) {
    linha++
  }
  var bebidas = produtos.getRange(2, 6, linha - 2, 1).getValues()
  linha = 2
  if (caixaat.getRange('C6').getValue() == 'Diurno') {
    while (produtos.getRange(linha, 7).isBlank() == false) {
      linha++
    }
    var refeicao = produtos.getRange(2, 7, linha - 2, 1).getValues()
  }
  else {
    while (produtos.getRange(linha, 9).isBlank() == false) {
      linha++
    }
    var refeicao = produtos.getRange(2, 9, linha - 2, 1).getValues()
  }
  linha = 2
  while (produtos.getRange(linha, 8).isBlank() == false) {
    linha++
  }
  var outros = produtos.getRange(2, 8, linha - 2, 1).getValues()
  linha = 2
  var Form = HtmlService.createTemplateFromFile("CaixaFuncionarios")
  Form.bebidas = bebidas.map(function (r) { return r[0]; });
  Form.refeicao = refeicao.map(function (r) { return r[0]; });
  Form.outros = outros.map(function (r) { return r[0]; });
  var Mostrar = Form.evaluate()
  Mostrar.setTitle("Registro de Comanda").setHeight(500).setWidth(800)
  SpreadsheetApp.getUi().showModalDialog(Mostrar, "Registro de Comanda")
}

function Vendas() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var produtos = spreadsheet.getSheetByName('Base de Preços')
  var linha = 1
  while (produtos.getRange(linha, 2).isBlank() == false) {
    linha = linha + 1
  }
  var list = produtos.getRange(2, 2, linha - 2, 1).getValues()
  var Form = HtmlService.createTemplateFromFile("Vendas")
  Logger.log(list)
  Form.list = list.map(function (r) { return r[0]; });
  var Mostrar = Form.evaluate()
  Mostrar.setTitle("Cadastro de Venda").setHeight(500).setWidth(400)
  SpreadsheetApp.getUi().showModalDialog(Mostrar, "Cadastro de Venda")
}

function enviar(Dados) {
  var spreadsheet = SpreadsheetApp.openById(id)
  var sheet = spreadsheet.getSheetByName('Fluxo de Caixa')
  sheet.getRange('J1').activate()
  sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
  sheet.getActiveCell().offset(1, 0).activate()
  var linha = sheet.getActiveCell().getRow()
  var data = new Date();
  var dia = String(data.getDate()).padStart(2, '0')
  var mes = String(data.getMonth() + 1).padStart(2, '0')
  var ano = data.getFullYear()
  var data = dia + '/' + mes + '/' + ano
  sheet.getRange(linha, 10).setValue(data)
  sheet.getRange(linha, 11).setValue([Dados.ent])
  sheet.getRange(linha, 12).setValue([Dados.forn])
  sheet.getRange(linha, 13).setValue([Dados.desc])
  if (Dados.ent == "Entrada"){
    sheet.getRange(linha, 14).setValue([Dados.val])
  }
  else{ 
    sheet.getRange(linha, 14).setValue([Dados.val]* (-1) )}
  sheet.getRange(linha, 15).setValue([Dados.obs])
  while (sheet.getRange(linha, 12).isBlank() == false) {
    linha = linha + 1
  }
  sheet.getRange(2, 10, linha - 2, 6).activate();
  sheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}

function Notas() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ativador = spreadsheet.getSheetByName('Base de preços')
  var linha = 2
  while (ativador.getRange(linha, 17).isBlank() == false) {
    linha++
  }
  var setores = ativador.getRange(2,19,(ativador.getRange(2,19).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()-1),1).getValues()
  var fornec = ativador.getRange(2, 17, linha - 2).getValues();
  var Form = HtmlService.createTemplateFromFile("Notas");
  Logger.log(fornec);
  Form.setores = setores.map(function (r) { return r[0]; });
  Form.fornec = fornec.map(function (r) { return r[0]; });
  var Mostrar = Form.evaluate()
  Mostrar.setTitle("Caixa").setHeight(500).setWidth(600)
  SpreadsheetApp.getUi().showModalDialog(Mostrar, "Caixa")
}

function enviar2(Dados) {
  var spreadsheet = SpreadsheetApp.openById(id)
  var sheet = spreadsheet.getSheetByName('Notas de Compra')
  linha = 2
  while (sheet.getRange(linha, 8).isBlank() == false) {
    linha = linha + 1
  }
  var data = new Date();
  sheet.getRange(linha, 8).setValue(data)
  sheet.getRange(linha, 9).setValue([Dados.ent])
  sheet.getRange(linha, 10).setValue([Dados.forn])
  sheet.getRange(linha, 11).setValue([Dados.desc])
  sheet.getRange(linha, 12).setValue([Dados.val * (-1)])
  sheet.getRange(linha, 13).setValue([Dados.obs])
  while (sheet.getRange(linha, 12).isBlank() == false) {
    linha = linha + 1
  }
  sheet.getRange(2, 8, linha - 2, 6).activate();
  sheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}



function onEditAction(e) {
  var ss = e.source.getActiveSheet()
  var range = e.range.getA1Notation()
  if (ss.getName() != 'Fluxo de Caixa' || range != 'G22') { return }
  Caixa()
  ss.getRange('G22').clearContent()
}
function enviarProduto(Dados) {
  var spreadsheet = SpreadsheetApp.openById(id)
  var sheet = spreadsheet.getSheetByName('Base de Preços')
  sheet.getRange('B1').activate()
  sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
  sheet.getActiveCell().offset(1, 0).activate()
  var linha = sheet.getActiveCell().getRow()
  sheet.getRange(linha, 1).setValue(linha - 1)
  sheet.getRange(linha, 2).setValue([Dados.nome])
  sheet.getRange(linha, 3).setValue([Dados.val])
  sheet.getRange(linha, 4).setValue([Dados.grupo])
  if (Dados.grupo == "Bebidas") {
    sheet.getRange('F2').activate()
    sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
    sheet.getActiveCell().offset(1, 0).activate()
    var linha = sheet.getActiveCell().getRow()
    sheet.getRange(linha, 6).setValue([Dados.nome])
  }
  if (Dados.grupo == "Restaurante") {
    sheet.getRange('G2').activate()
    sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
    sheet.getActiveCell().offset(1, 0).activate()
    var linha = sheet.getActiveCell().getRow()
    sheet.getRange(linha, 7).setValue([Dados.nome])
  }
  if (Dados.grupo == "Diversos") {
    sheet.getRange('H2').activate()
    sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
    sheet.getActiveCell().offset(1, 0).activate()
    var linha = sheet.getActiveCell().getRow()
    sheet.getRange(linha, 8).setValue([Dados.nome])
  }
  if (Dados.grupo == "Delivery") {
    sheet.getRange('I2').activate()
    sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
    sheet.getActiveCell().offset(1, 0).activate()
    var linha = sheet.getActiveCell().getRow()
    sheet.getRange(linha, 9).setValue([Dados.nome])
  }
  sheet.getRange('F:I').setHorizontalAlignment("center")
}

function enviarVenda(Dados) {
  var spreadsheet = SpreadsheetApp.openById(id)
  var sheet = spreadsheet.getSheetByName('Vendas')
  var data = new Date();
  var dia = String(data.getDate()).padStart(2, '0')
  var mes = String(data.getMonth() + 1).padStart(2, '0')
  var ano = data.getFullYear()
  var data = dia + '/' + mes + '/' + ano
  var hora = data.getHours
  if (hora <= 15) {
    var turno = 'Diurno'
  } else {
    var turno = 'Noturno'
  }
  Logger.log(turno)
  var produtos = spreadsheet.getSheetByName('Base de Preços')
  lin = 1
  while (produtos.getRange(lin, 2).isBlank() == false) {
    lin = lin + 1
  }
  // Pegando uma lista com os nomes dos produtos e seus valores
  var list = produtos.getRange(2, 2, lin - 2, 1).getValues()
  var valor = produtos.getRange(2, 3, lin - 2, 1).getValues()
  for (var i = 0; i < list.length; i++) {
    for (var j = 0; j < list[i].length; j++) {
      if (list[i][j] == Dados.prod) {
        var val = valor[i][j]
      }
    }
  }
  // Salvando dados nas células
  sheet.getRange('A1').activate()
  sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
  sheet.getActiveCell().offset(1, 0).activate()
  var linha = sheet.getActiveCell().getRow()
  sheet.getRange(linha, 1).setValue(data)
  sheet.getRange(linha, 2).setValue(turno)
  sheet.getRange(linha, 3).setValue([Dados.prod])
  sheet.getRange(linha, 4).setValue([Dados.qnt])
  sheet.getRange(linha, 5).setValue(val * [Dados.qnt])
}

function pesquisa(DadosP) {
  var data = DadosP.data
  var splitdata = data.split('-')
  var refdata = (splitdata[2] + '/' + splitdata[1] + '/' + splitdata[0])
  var sheet = SpreadsheetApp.getActiveSheet();
  var linha = 2
  while (sheet.getRange(linha, 10).getValue() != refdata) {
    linha++
  }
  var seta = linha
  while (sheet.getRange(linha, 10).getValue() == refdata) {
    linha++
  }
  sheet.getRange(seta, 10, linha - seta, 5).activate().setBackground('#80b388').activate()
};

function removerfiltro() {
  sheet = SpreadsheetApp.getActiveSheet()
  var linha = 2
  Logger.log(sheet.getRange(linha, 11).getBackground())
  while (sheet.getRange(linha, 10).getBackground() == '#ffffff') {
    linha++
  }
  Logger.log(linha)
  var seta = linha
  while (sheet.getRange(linha, 10).getBackground() != '#ffffff') {
    linha++
  }
  Logger.log(linha)
  sheet.getRange(seta, 10, linha - seta, 5).setBackground('white')
};

function Aporte() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Tela do Caixa'), true);
  var Form = HtmlService.createTemplateFromFile("Aporte")
  var Mostrar = Form.evaluate()
  Mostrar.setTitle("Aporte").setHeight(300).setWidth(250)
  SpreadsheetApp.getUi().showModalDialog(Mostrar, "Aporte")
}
function entradacomanda(products, quanti, formapg, tipovenda) {
  var spreadsheet = SpreadsheetApp.openById(id)
  var sheet = spreadsheet.getSheetByName('Tela do Caixa')
  var produtos = spreadsheet.getSheetByName('Base de Preços')
  var lin = produtos.getLastRow()
  var valores = produtos.getRange(2, 2, lin - 1, 3).getValues()
  sheet.getRange('B9').activate()
  sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
  sheet.getActiveCell().offset(1, 0).activate()
  var linha = sheet.getActiveCell().getRow()
  var seta = linha
  sheet.getRange(linha, 1).setValue('C' + seta)
  sheet.getRange(linha, 5).setValue(formapg)
  sheet.getRange(linha, 6).setValue(tipovenda)
  var comandav = 0
  for (i = 0; i < products.length; i++) {
    sheet.getRange(linha, 2).setValue(products[i])
    sheet.getRange(linha, 3).setValue(quanti[i])
    for (j = 0; j < valores.length; j++) {
      if (products[i] == valores[j][0]) {
        sheet.getRange(linha, 4).setValue(valores[j][1] * quanti[i])
        var comandav = comandav + valores[j][1] * quanti[i]
      }
    }
    linha = linha + 1
  }
  if (tipovenda == "Salão") {
    if (formapg == "Dinheiro") {
      var ant = sheet.getRange('J31').getValue()
      var novo = comandav
      sheet.getRange('J31').setValue(novo + ant)
    }
    if (formapg == "Crédito") {
      var ant = sheet.getRange('J32').getValue()
      var novo = comandav - (comandav * (produtos.getRange('V5').getValue()))
      sheet.getRange('J32').setValue(novo + ant)
    }
    if (formapg == "Débito") {
      var ant = sheet.getRange('J33').getValue()
      var novo = comandav - comandav * (produtos.getRange('V5').getValue())
      sheet.getRange('J33').setValue(novo + ant)
    }
    if (formapg == "Pix") {
      var ant = sheet.getRange('J34').getValue()
      var novo = comandav
      sheet.getRange('J34').setValue(novo + ant)
    }
  }
  if (tipovenda == "Delivery") {
    if (formapg == "Dinheiro") {
      var ant = sheet.getRange('J43').getValue()
      var novo = comandav
      sheet.getRange('J43').setValue(novo + ant)
    }
    if (formapg == "Crédito") {
      var ant = sheet.getRange('J45').getValue()
      var novo = comandav - (comandav * (produtos.getRange('V5').getValue()))
      sheet.getRange('J45').setValue(novo + ant)
    }
    if (formapg == "Débito") {
      var ant = sheet.getRange('J45').getValue()
      var novo = comandav - (comandav * (produtos.getRange('V5').getValue()))
      sheet.getRange('J45').setValue(novo + ant)
    }
    if (formapg == "Pix") {
      var ant = sheet.getRange('J44').getValue()
      var novo = comandav
      sheet.getRange('J44').setValue(novo + ant)
    }
  }
  if (tipovenda == "Ifood") {
    if (formapg == "Dinheiro") {
      var ant = sheet.getRange('J49').getValue()
      var novo = comandav 
      sheet.getRange('J49').setValue(novo + ant)
    }
    if (formapg == "Crédito") {
      var ant = sheet.getRange('J50').getValue()
      var novo = comandav  - (comandav * (produtos.getRange('V5').getValue()))
      sheet.getRange('J50').setValue(novo + ant)
    }
    if (formapg == "Débito") {
      var ant = sheet.getRange('J50').getValue()
      var novo = comandav  - (comandav * (produtos.getRange('V5').getValue()))
      sheet.getRange('J50').setValue(novo + ant)
    }
    if (formapg == "Pix") {
      var ant = sheet.getRange('J51').getValue()
      var novo = comandav 
      sheet.getRange('J51').setValue(novo + ant)
    }
    if (formapg == "Aplicativo") {
      var ant = sheet.getRange('J52').getValue()
      var novo = comandav - (comandav * (produtos.getRange('V2').getValue()))
      sheet.getRange('J52').setValue(novo + ant)
    }
  }
  if (tipovenda == "99 Food") {
    if (formapg == "Dinheiro") {
      var ant = sheet.getRange('J56').getValue()
      var novo = comandav
      sheet.getRange('J56').setValue(novo + ant)
    }
    if (formapg == "Crédito") {
      var ant = sheet.getRange('J57').getValue()
      var novo = comandav  - (comandav * (produtos.getRange('V5').getValue()))
      sheet.getRange('J57').setValue(novo + ant)
    }
    if (formapg == "Débito") {
      var ant = sheet.getRange('J57').getValue()
      var novo = comandav  - (comandav * (produtos.getRange('V5').getValue()))
      sheet.getRange('J57').setValue(novo + ant)
    }
    if (formapg == "Aplicativo") {
      var ant = sheet.getRange('J59').getValue()
      var novo = comandav  - (comandav * (produtos.getRange('V4').getValue()))
      sheet.getRange('J59').setValue(novo + ant)
    }
    if (formapg == "Pix") {
      var ant = sheet.getRange('J58').getValue()
      var novo = comandav
      sheet.getRange('J58').setValue(novo + ant)
    }
  }
  if (tipovenda == "Aiqfome") {
    if (formapg == "Dinheiro") {
      var ant = sheet.getRange('J63').getValue()
      var novo = comandav
      sheet.getRange('J63').setValue(novo + ant)
    }
    if (formapg == "Crédito") {
      var ant = sheet.getRange('J64').getValue()
      var novo = comandav  - (comandav * (produtos.getRange('V5').getValue())) 
      sheet.getRange('J64').setValue(novo + ant)
    }
    if (formapg == "Débito") {
      var ant = sheet.getRange('J64').getValue()
      var novo = comandav  - (comandav * (produtos.getRange('V5').getValue()))
      sheet.getRange('J64').setValue(novo + ant)
    }
    if (formapg == "Aplicativo") {
      var ant = sheet.getRange('J66').getValue()
      var novo = comandav  - (comandav * (produtos.getRange('V3').getValue()))
      sheet.getRange('J66').setValue(novo + ant)
    }
    if (formapg == "Pix") {
      var ant = sheet.getRange('J65').getValue()
      var novo = comandav
      sheet.getRange('J65').setValue(novo + ant)
    }
  }

  sheet.getRange(seta, 1, products.length, 1).activate().mergeVertically();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(seta, 5, products.length, 2).activate().mergeVertically();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center').setVerticalAlignment('middle');
  var linha = sheet.getRange('B9').getRow()
  while (sheet.getRange(linha, 2).isBlank() == false) {
    linha = linha + 1
  }
  sheet.getRange(11, 1, linha - 11, 6).activate();
  sheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}
function Save(Dados) {
  var spreadsheet = SpreadsheetApp.openById(id)
  var sheet = spreadsheet.getSheetByName('Tela do Caixa')
  sheet.getRange('B9').activate()
  sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
  sheet.getActiveCell().offset(1, 0).activate()
  var linha = sheet.getActiveCell().getRow()
  sheet.getRange(linha, 1).setValue('AP' + linha).setHorizontalAlignment('center')
  sheet.getRange(linha, 4).setValue([Dados.Quantidade])
  sheet.getRange(linha, 2).setValue([Dados.Desc])
  sheet.getRange(linha, 6).setValue("Aporte").setHorizontalAlignment('center')
  sheet.getRange(linha, 5).setValue("Aporte").setHorizontalAlignment('center')
  sheet.getRange(linha, 3).setValue("AP")
  var ant = sheet.getRange('J29').getValue()
  var novo = sheet.getRange(linha, 4).getValue()
  sheet.getRange('J29').setValue(novo + ant)
  while (sheet.getRange(linha, 2).isBlank() == false) {
    linha = linha + 1
  }
  sheet.getRange(11, 1, linha - 11, 6).activate();
  sheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}

function Sangria(Dados) {
  var spreadsheet = SpreadsheetApp.openById(id)
  var sheet = spreadsheet.getSheetByName('Tela do Caixa')
  sheet.getRange('B9').activate()
  sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate()
  sheet.getActiveCell().offset(1, 0).activate()
  var linha = sheet.getActiveCell().getRow()
  sheet.getRange(linha, 1).setValue('SG' + linha).setHorizontalAlignment('center')
  sheet.getRange(linha, 4).setValue([Dados.Quantidade])
  sheet.getRange(linha, 2).setValue([Dados.Desc])
  sheet.getRange(linha, 5).setValue([Dados.Nome]).setHorizontalAlignment('center')
  sheet.getRange(linha, 6).setValue("Sangria").setHorizontalAlignment('center')
  sheet.getRange(linha, 3).setValue("SG")
  var ant = sheet.getRange('J30').getValue()
  var novo = sheet.getRange(linha, 4).getValue()
  sheet.getRange('J30').setValue(novo + ant)
  while (sheet.getRange(linha, 2).isBlank() == false) {
    linha = linha + 1
  }
  sheet.getRange(11, 1, linha - 11, 6).activate();
  sheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}

function Fechat() {
  var spreadsheet = SpreadsheetApp.openById(id)
  var sheet = spreadsheet.getSheetByName('Base de Preços')
  var guia = spreadsheet.getSheetByName('Dados de Guias')
  var notas = spreadsheet.getSheetByName('Notas de Compra')
  var fluxo = spreadsheet.getSheetByName('Fluxo de Caixa')
  var beb = sheet.getRange(2,6,sheet.getRange(2, 6).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()-1,1).getValues()
  sheet = spreadsheet.getSheetByName('Tela do Caixa')
  sheet.getRange('C8').setValue(sheet.getRange('I4').getValue())
  var linha = 11
  var seta = linha
  var line = 2
  var marca = fluxo.getRange(fluxo.getRange(2,10).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1,0).getRow(),10)
  while (notas.getRange(line, 8).isBlank() == false){
    line++}
  notas.getRange(2,8,line-2,6).moveTo(marca)

  while (sheet.getRange(linha, 2).isBlank() == false) {
    linha = linha + 1}
  
  var outro = spreadsheet.getSheetByName('Dados')
  var ahnil = 2
  while (outro.getRange(ahnil, 13).isBlank() == false) {
    ahnil = ahnil + 1
  }
  var point = outro.getRange(ahnil, 13)
  var dia = sheet.getRange('C4').getValue()
  var turno = sheet.getRange('C6').getValue()
  var digital = sheet.getRange('I9').getValue()
  var Dados = {ent: "Entrada",forn: "Fechamento",desc: turno, val: digital}
  enviar(Dados) 
  outro.getRange(outro.getRange(1,28).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1,0).getRow(),28).setValue(dia)
  outro.getRange(outro.getRange(1,29).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1,0).getRow(),29).setValue(turno)
  outro.getRange(outro.getRange(1,30).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1,0).getRow(),30).setValue(sheet.getRange('I4').getValue())
     
  outro.getRange(outro.getRange(1,32).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1,0).getRow(),32).setValue(dia)
  outro.getRange(outro.getRange(1,33).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1,0).getRow(),33).setValue(turno)
  outro.getRange(outro.getRange(1,34).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1,0).getRow(),34).setValue(sheet.getRange('I9').getValue())
  if (turno == 'Diurno') {
    sheet.getRange('C6').setValue("Noturno")
  }
  else {
    sheet.getRange('C6').setValue("Diurno")
  }

  outro.getRange(1, 30).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0).setValue(sheet.getRange('J10').getValue())
  outro.getRange(ahnil, 11).setValue(turno)
  outro.getRange(ahnil, 10).setValue(dia)
  sheet.getRange(seta, 2, linha - seta, 5).moveTo(point)
  sheet.getRange('A11:A').deleteCells(SpreadsheetApp.Dimension.ROWS)
  sheet.getRange(31, 9, 4, 2).copyTo(guia.getRange(2, 3).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0))
  guia.getRange(guia.getRange(2, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0).getRow(), 1, 4, 1).setValue(dia)
  guia.getRange(guia.getRange(2, 2).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0).getRow(), 2, 4, 1).setValue(turno)
  sheet.getRange(43, 9, 3, 2).copyTo(guia.getRange(2, 8).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0))
  guia.getRange(guia.getRange(2, 6).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0).getRow(), 6, 3, 1).setValue(dia)
  guia.getRange(guia.getRange(2, 7).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0).getRow(), 7, 3, 1).setValue(turno)
  sheet.getRange(49, 9, 4, 2).copyTo(guia.getRange(2, 13).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0))
  guia.getRange(guia.getRange(2, 11).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0).getRow(), 11, 4, 1).setValue(dia)
  guia.getRange(guia.getRange(2, 12).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0).getRow(), 12, 4, 1).setValue(turno)
  sheet.getRange(56, 9, 4, 2).copyTo(guia.getRange(2, 18).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0))
  guia.getRange(guia.getRange(2, 16).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0).getRow(), 16, 4, 1).setValue(dia)
  guia.getRange(guia.getRange(2, 17).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0).getRow(), 17, 4, 1).setValue(turno)
  sheet.getRange(63, 9, 4, 2).copyTo(guia.getRange(2, 23).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0))
  guia.getRange(guia.getRange(2, 21).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0).getRow(), 21, 4, 1).setValue(dia)
  guia.getRange(guia.getRange(2, 22).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0).getRow(), 22, 4, 1).setValue(turno)

  sheet.getRange(29, 10, 6, 1).setValue(0)
  sheet.getRange(43, 10, 3, 1).setValue(0)
  sheet.getRange(49, 10, 4, 1).setValue(0)
  sheet.getRange(56, 10, 4, 1).setValue(0)
  sheet.getRange(63, 10, 4, 1).setValue(0)
  outro.getRange('P:Q').breakApart()
  
  
  var i = ahnil
  outro.getRange(i, 12).setFormula("=VLOOKUP(M" + i + ";'Base de Preços'!$B:$D;3;)")
  for (i = ahnil + 1; i < linha - seta + ahnil; i++) {
    outro.getRange(i, 12).setFormula("=VLOOKUP(M" + i + ";'Base de Preços'!$B:$D;3;)")
    outro.getRange(i, 11).setValue(turno)
    outro.getRange(i, 10).setValue(dia)
    if(outro.getRange(i,12).getValue()== 'Bebidas'){
      for(j=0;j<beb.length;j++){
        if(outro.getRange(i,13).getValue()==beb[j][0]){
          qnt = outro.getRange(i,14).getValue()
          forma(qnt,outro.getRange(i,13).getValue())
        }
      }
    }
    if (outro.getRange(i, 17).isBlank() == true) {
      outro.getRange(i, 17).setValue(outro.getRange(i - 1, 17).getValue())
    }
    if (outro.getRange(i, 16).isBlank() == true) {
      outro.getRange(i, 16).setValue(outro.getRange(i - 1, 16).getValue())
    }
  }
  var h = outro.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0).getRow()
  var g = outro.getRange(1, 19).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0).getRow()
  j = ahnil
  while (outro.getRange(j, 17).isBlank() != true) {
    if (outro.getRange(j, 17).getValue() == "Aporte") {
      var pt = outro.getRange(g, 19)
      outro.getRange(j, 10, 1, 8).moveTo(pt)
      outro.getRange(g, 21).setValue("Aporte")
      outro.getRange(j, 10, 1, 8).deleteCells(SpreadsheetApp.Dimension.ROWS)
      j--
      g++
    }
    if (outro.getRange(j, 17).getValue() == "Sangria") {
      var pt = outro.getRange(h, 1)
      outro.getRange(j, 10, 1, 8).moveTo(pt)
      outro.getRange(h, 3).setValue("Sangria")
      outro.getRange(j, 10, 1, 8).deleteCells(SpreadsheetApp.Dimension.ROWS)
      j--
      h++
    }
    j++
  }
  outro.getRange(2, 1, h - 2, 8).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
  outro.getRange(2, 10, j - 2, 8).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
  outro.getRange(2, 19, g - 2, 8).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
}
function forma(qnt, beb) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Freezers')
  var f1 = sheet.getRange("D3:D8").getValues()
  var t1 = sheet.getRange("B3:B8").getValues()
  var f2 = sheet.getRange("D12:D16").getValues()
  var t2 = sheet.getRange("B12:B16").getValues()
  var f3 = sheet.getRange("D20:D27").getValues()
  var t3 = sheet.getRange("B20:B27").getValues()
  var f4 = sheet.getRange("D31:D34").getValues()
  var t4 = sheet.getRange("B31:B34").getValues()
  var f5 = sheet.getRange("D38:D43").getValues()
  var t5 = sheet.getRange("B38:B43").getValues()
  for (i = 0; i < f1.length; i++) {
    if(sheet.getRange(i+3,3).getValue()==beb){
      sheet.getRange(i+3,4).setValue(f1[i][0]-qnt)
    }
    if (f1[i][0] >= t1[i][0] * 0.75) {
      sheet.getRange(i + 3, 4).setBackground('#d9ead3')
    }
    if (f1[i][0] > t1[i][0] * 0.25 && f1[i][0] < t1[i][0] * 0.75) {
      sheet.getRange(i + 3, 4).setBackground('#fff2cc')
    }
    if (f1[i][0] <= t1[i][0] * 0.25) {
      sheet.getRange(i + 3, 4).setBackground('#f4c7c3')
    }
  }
  for (i = 0; i < f2.length; i++) {
    if(sheet.getRange(i+12,3).getValue()==beb){
      sheet.getRange(i+12,4).setValue(f2[i][0]-qnt)
    }
    if (f2[i][0] >= t2[i][0] * 0.75) {
      sheet.getRange(i + 12, 4).setBackground('#d9ead3')
    }
    if (f2[i][0] > t2[i][0] * 0.25 && f2[i][0] < t2[i][0] * 0.75) {
      sheet.getRange(i + 12, 4).setBackground('#fff2cc')
    }
    if (f2[i][0] <= t2[i][0] * 0.25) {
      sheet.getRange(i + 12, 4).setBackground('#f4c7c3')
    }
  }
  for (i = 0; i < f3.length; i++) {
    if(sheet.getRange(i+20,3).getValue()==beb){
      sheet.getRange(i+20,4).setValue(f3[i][0]-qnt)
    }
    if (f3[i][0] >= t3[i][0] * 0.75) {
      sheet.getRange(i + 20, 4).setBackground('#d9ead3')
    }
    if (f3[i][0] > t3[i][0] * 0.25 && f3[i][0] < t3[i][0] * 0.75) {
      sheet.getRange(i + 20, 4).setBackground('#fff2cc')
    }
    if (f3[i][0] <= t3[i][0] * 0.25) {
      sheet.getRange(i + 20, 4).setBackground('#f4c7c3')
    }
  }
  for (i = 0; i < f4.length; i++) {
    if(sheet.getRange(i+31,3).getValue()==beb){
      sheet.getRange(i+31,4).setValue(f4[i][0]-qnt)
    }
    if (f4[i][0] >= t4[i][0] * 0.75) {
      sheet.getRange(i + 31, 4).setBackground('#d9ead3')
    }
    if (f4[i][0] > t4[i][0] * 0.25 && f4[i][0] < t4[i][0] * 0.75) {
      sheet.getRange(i + 31, 4).setBackground('#fff2cc')
    }
    if (f4[i][0] <= t4[i][0] * 0.25) {
      sheet.getRange(i + 31, 4).setBackground('#f4c7c3')
    };
  }
  for (i = 0; i < f5.length; i++) {
    if(sheet.getRange(i+38,3).getValue()==beb){
      sheet.getRange(i+38,4).setValue(f5[i][0]-qnt)
    }
    if (f5[i][0] >= t5[i][0] * 0.75) {
      sheet.getRange(i + 38, 4).setBackground('#d9ead3')
    }
    if (f5[i][0] > t5[i][0] * 0.25 && f5[i][0] < t5[i][0] * 0.75) {
      sheet.getRange(i + 38, 4).setBackground('#fff2cc')
    }
    if (f5[i][0] <= t5[i][0] * 0.25) {
      sheet.getRange(i + 38, 4).setBackground('#f4c7c3')
    }
  }
}

