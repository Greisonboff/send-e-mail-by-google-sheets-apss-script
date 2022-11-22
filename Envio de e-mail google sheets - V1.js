function aplication() {

    //Hora e data
    var dataAtual = new Date();
    var dia = dataAtual.getDate();
    var mes = (dataAtual.getMonth() + 1);
    var ano = dataAtual.getFullYear();
    var date = dia + '/' + mes + '/' + ano;

    var diaSemana = dataAtual.getDay()// Dia da semana '0 para Domingo'
    var horas = dataAtual.getHours();
    var minutos = dataAtual.getMinutes();
    var horaAtual = horas + ":" + minutos;

    // validate table by time
    if (horas == 9 || horas == 09) {
        // table name
        var nomeTabela = 'Events';
        
        // table columns
        var quantitiEvents = 'I2';
        var category = 'A';
        var action = 'B';
        var label = 'C';
        var totalEvents = 'E';
        var uniqueEvents = 'F';
        var validationQuantitiEvents = 20;
        // table times
        var timeLastRunReports = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeTabela).getRange('I6').getValue();
        var lastDay = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeTabela).getRange('D2').getValue();
    } else {
        // table name
        var nomeTabela = 'BQ';
        
        // table columns
        var quantitiEvents = 'I2';
        var category = 'A';
        var action = 'B';
        var label = 'C';
        var totalEvents = 'E';
        var uniqueEvents = 'D';
        var validationQuantitiEvents = 0;
        // table times
        var timeLastRunReports = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeTabela).getRange('I6').getValue();
        var lastDay = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeTabela).getRange('I6').getValue();
    }


    if (horas == 9 || horas == 09 || horas == 14) {
        // Add numero de mensagens
        var quantitiEventsValue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeTabela).getRange(quantitiEvents).getValue();
        quantitiEventsValue = quantitiEventsValue + 2;

        var information = [];
        var informationDis = [];
        var informationPos = [];
        var informationChk = [];

        for (var i = 2; i < quantitiEventsValue; ++i) {
            var actions = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeTabela).getRange(action + i).getValue().replaceAll('[', '').replaceAll(']', '').split('|');
            if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeTabela).getRange(uniqueEvents + i).getValue() > validationQuantitiEvents) {
                var dataRange = '<p><b class="title">Pagina: </b>' + actions[0] +
                    '<b class="title marginBold">  Tribo: </b>' + actions[1] +
                    '<b class="title marginBold">  Nº teste: </b>' + actions[2] +
                    '<b class="title marginBold">  Variante: </b>' + actions[3] +
                    '<b class="title marginBold">  Eventos</b> Total: <b class="title marginBold">' + SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeTabela).getRange(totalEvents + i).getValue() +
                    '</b> Exclusivos: <b class="title marginBold">' + SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeTabela).getRange(uniqueEvents + i).getValue() +
                    '</b><br> Erro: <b class="title">' + SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeTabela).getRange(label + i).getValue() + '</b><br></p>';
                if (actions.join('').indexOf('DIS') != -1) {
                    informationDis.push(dataRange);
                } else if (actions.join('').indexOf('POS') != -1) {
                    informationPos.push(dataRange);
                } else if (actions.join('').indexOf('CHE') != -1) {
                    informationChk.push(dataRange);
                }
            }
        }

        // break apart tribo
        if (informationDis.length > 0) {
            informationDis = (informationDis.join(''));
            informationDis = '<tr><td class="conteudo borderTd"><h3>Tribo Discovery</h3>' + informationDis + '</td></tr><tr><td><p> </p></td></tr>';
            information.push(informationDis);
        }

        if (informationPos.length > 0) {
            informationPos = (informationPos.join(''));
            informationPos = '<tr><td class="conteudo borderTd"><h3>Tribo Pós venda</h3>' + informationPos + '</td></tr><tr><td><p> </p></td></tr>';
            information.push(informationPos);
        }

        if (informationChk.length > 0) {
            informationChk = (informationChk.join(''));
            informationChk = '<tr><td class="conteudo borderTd"><h3>Tribo Checkout</h3>' + informationChk + '</td></tr><tr><td><p> </p></td></tr>';
            information.push(informationChk);
        }

        console.log(information)

        //Hora e data
        var dataAtualReport = new Date(lastDay);
        var diaReport = dataAtualReport.getDate();
        var mesReport = (dataAtualReport.getMonth() + 1);
        var anoReport = dataAtualReport.getFullYear();
        var dateReport = diaReport + '/' + mesReport + '/' + anoReport;

        var diaSemana = dataAtualReport.getDay()// Dia da semana '0 para Domingo'
        var horasReport = dataAtualReport.getHours();
        var minutosReport = dataAtualReport.getMinutes();
        var horaAtualReport = horasReport + ":" + minutosReport;

        information = (information.join(''));
        ////Mensagens de erro script
        function getMessage() {
            var htmlOutput = HtmlService.createHtmlOutputFromFile('index');
            var message = htmlOutput.getContent();
            message = message.replace("%tabelaGeral", information);
            message = message.replace("%data", dateReport);
            return message;
        }


        // email config
        if (information.length > 0) {
            // Set email address
            var contEmail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("E-mail").getRange("C1").getValues();
            var email;
            var posEmail = 2;

            for (var i = 0; i < contEmail; ++i) {
                var email = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("E-mail").getRange("B" + posEmail).getValues();
                var emailAddress = '' + email;
                var message = getMessage(information);
                var subject = "Renner Erros teste AB";
                console.log(message)
                MailApp.sendEmail(emailAddress, subject, message, { htmlBody: message });
                posEmail++;
            }
        }
    }
}