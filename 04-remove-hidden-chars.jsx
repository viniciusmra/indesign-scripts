/*
    Autor: Vinicius Alves - viniciusmra@gmail.com
    Data: set/2021
    Última atualização: jun/2022
    
	Descrição:
    O script exclui ou substitui alguns erros de formatação que por ventura venham com o documento:
    - Vários espaços seguidos
    - Vários enters seguidos
    - Espaço + enter
    - Enter + espaço
    - Nonbreaking spaces (Espaços que não podem ser quebrados em linhas)
    - Forced linebreaks (Quebra de linha forçada)
    - Tabs (Algumas pessoas usam tabs para formatar os documentos)

    OBS: Talvez seja preciso rodar o script mais de uma vez

*/

//Limpa as preferências do Grep
app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;

//Configura a busca
app.findChangeGrepOptions.includeFootnotes = true;
app.findChangeGrepOptions.includeHiddenLayers = true;
app.findChangeGrepOptions.includeLockedLayersForFind = false;
app.findChangeGrepOptions.includeLockedStoriesForFind = false;
app.findChangeGrepOptions.includeMasterPages = false;

//Limpa as preferências do texto
app.findTextPreferences = NothingEnum.nothing;
app.changeTextPreferences = NothingEnum.nothing;

//Configura a busca
app.findChangeTextOptions.includeFootnotes = true;


function dialog1(){
	var myWindow = new Window ("dialog", "Encontrados");

    // Texto
	myWindow.add("statictext", undefined, "Foram encontrados: ").alignment = 'left'
	
    myWindow.add("statictext", undefined, "Espaços seguidos: " + a).alignment = 'left'
	myWindow.add("statictext", undefined, "Espaço + enter: " + b).alignment = 'left'
	myWindow.add("statictext", undefined, "Enter + espaço: " + c).alignment = 'left'
	myWindow.add("statictext", undefined, "Enters seguidos: " + d).alignment = 'left'
    myWindow.add("statictext", undefined, "Nonbreaking spaces: " + e).alignment = 'left'
    myWindow.add("statictext", undefined, "forced linebreaks: " + f).alignment = 'left'
    myWindow.add("statictext", undefined, "Tabs: " + g).alignment = 'left'
	
    // Grupo 1
	var group1 = myWindow.add('group {alignment:["center","fill"]}')
	group1.orientation = "row";
	group1.add("button", undefined, "Remover caracteres", {name:"ok"})
	group1.add("button", undefined, "Cancelar", {name:"cancel"})
	
	return myWindow.show();
}

a = changeGREP("\\x{20}\\x{20}+", "\\x{20}", false); // Substitui espaços seguidos pelo espaço único

b = changeGREP("\\x{20}\\r", "\\r", false); // Substitui um espaço seguido de enter pelo enter

c = changeGREP("\\r\\x{20}", "\\r", false); // Substitui um enter seguido de espaço (espaço no começo do paragrafo)

d = changeGREP("\\r\\r+", "\\r", false); // Substitui vários enter seguidos

e = changeGREP("~S", "\\s", false); // Subistitui non breaking space (^) por um espaço simples

f = changeText("^n", "^p", false); // Subistiui o forced line break por um return

g = changeGREP("\\t", "", false); // Remove os tabs

if(dialog1() == 1){
    //Busca e substituição
    changeGREP("\\x{20}\\x{20}+", "\\x{20}", true); // Substitui espaços seguidos pelo espaço único

    changeGREP("\\t", "", true); // Remove os tabs

    changeGREP("\\x{20}\\r", "\\r", true); // Substitui um espaço seguido de enter pelo enter

    changeGREP("\\r\\x{20}", "\\r", true); // Substitui um enter seguido de espaço (espaço no começo do paragrafo)

    changeGREP("\\r\\r+", "\\r", true); // Substitui vários enter seguidos

    changeGREP("~S", "\\s", true); // Subistitui non breaking space (^) por um espaço simples

    changeText("^n", "^p", true); // Subistiui o forced line break por um return

    //Limpa as preferencias de Grep
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;

    //Limpa as preferencias de texto
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;
}

// GREP
function changeGREP(find, change, applyChange){
    app.findGrepPreferences.findWhat = find; //procura
    if(applyChange){
        app.changeGrepPreferences.changeTo = change; // substitui
        app.changeGrep(); //conclui a substituição
    } else{
        return app.activeDocument.findGrep().length
    }
}
// Texto
function changeText(find, change, applyChange){
    app.findTextPreferences.findWhat = find; //procura
    if(applyChange){
        app.changeTextPreferences.changeTo = change; //substitui
        app.changeText(); //conclui a substituição
    } else {
        return app.activeDocument.findText().length
    }
}