/*
	Autor: Vinicius Alves - viniciusmra@gmail.com
    Data: set/2021
    Última atualização: jun/2022
    
	Descrição:  
    Este script tem a função de procurar preferencias de texto (negrito, italico, negrito & itálico e sobrescrito) dentro do documento
	e aplicar estilos de caracteres correspondentes, que podem ser escolhidos pelo usuário.
	Importante ressaltar que preferencias de texto são formatações vindas do editor de texto (Word, por exemplo), no entando elas podem ser
	sobrescritas por estilos de parágrafos (paragraph Styles), além disso não existe a possibilidade de formatar todas elas de uma vez só
	Para isso é impressindível aplicar estilos de caracteres (characterStyles).

*/
currentDoc = app.activeDocument;
docCharStyles = currentDoc.characterStyles.everyItem().name // lista de estilos de caracteres do documento.


textPreferences = ["Bold", "Italic", "Bold Italic", Position.superscript]; 	// Lista com as preferencias de texto
charStyleNames = [];														// Lista com o nome dos estilos de caractere que estão no documento

// Limpa os parametros de busca de trechos
function clearSettings(){
	app.findTextPreferences = NothingEnum.nothing;
	app.changeTextPreferences = NothingEnum.nothing;
	app.findChangeTextOptions.includeFootnotes = true;
}

// Primeira janela
// Configura qual os estilos de caracteres serão associdados a cada preferencia de texto
function dialog1(){
	var myWindow = new Window ("dialog", "Aplicar estilo de caractere");
	//myWindow.orientation = "column";

	//var group0 = window.add('group {alignment:["left","fill"]}')
	myWindow.add ("statictext", undefined, "Escolhas os estilos de caracteres:").alignment = "left";

	// Grupo 1
	var group1 = myWindow.add('group {alignment:["right","fill"]}')
	group1.orientation = "row";
	group1.add ("statictext", undefined, "Negrito:");
	var dp1 = group1.add ("dropdownlist", undefined, docCharStyles);
	dp1.selection = 2;

	// Grupo 2
	var group2 = myWindow.add('group {alignment:["right","fill"]}')
	group2.orientation = "row";
	group2.add ("statictext", undefined, "Itálico:");
	var dp2 = group2.add ("dropdownlist", undefined, docCharStyles);
	dp2.selection = 1;

	// Grupo 3
	var group3 = myWindow.add('group {alignment:["right","fill"]}')
	group3.orientation = "row";
	group3.add ("statictext", undefined, "Negrito & Itálico:");
	var dp3 = group3.add ("dropdownlist", undefined, docCharStyles);
	dp3.selection = 3;

	// Grupo 4
	var group4 = myWindow.add('group {alignment:["right","fill"]}')
	group4.orientation = "row";
	group4.add ("statictext", undefined, "Sobrescrito:");
	var dp4 = group4.add ("dropdownlist", undefined, docCharStyles);
	dp4.selection = 5;

	// Grupo 5
	var group5 = myWindow.add('group {alignment:["center","fill"]}')
	group5.orientation = "row";
	group5.add("button", undefined, "Procurar estilos", {name:"ok"})
	group5.add("button", undefined, "Cancelar", {name:"cancel"})
	bold = dp1.selection.text;

	if (myWindow.show() == 1){
		charStyleNames.push(dp1.selection.text);
		charStyleNames.push(dp2.selection.text);
		charStyleNames.push(dp3.selection.text);
		charStyleNames.push(dp4.selection.text);
		return 1;
	} else{
		return 0;
	}
}

// Segunda janela
// Mostra quantos trechos foram encontrados na busca
function dialog2(nBold, nItalic, nBoldItalic, nSuper){
	var myWindow = new Window ("dialog", "Encontrados");

	myWindow.add("statictext", undefined, "Foram encontrados: ").alignment = 'left'
	
	// Grupo 1
	var group1 = myWindow.add('group {alignment:["fill","fill"]}')
	group1.orientation = "column";
	var text1 = group1.add("statictext", undefined, "Negrito: " + nBold);
	text1.alignment = 'left'
	var text2 = group1.add("statictext", undefined, "Itálico: " + nItalic);
	text2.alignment = 'left'
	var text3 = group1.add("statictext", undefined, "Negrito Itálico: " + nBoldItalic);
	text3.alignment = 'left'
	var text4 = group1.add("statictext", undefined, "Sobrescrito: " + nSuper);
	text4.alignment = 'left'
	
	// Grupo 2
	var group2 = myWindow.add('group {alignment:["center","fill"]}')
	group2.orientation = "row";
	group2.add("button", undefined, "Aplicar estilos", {name:"ok"})
	group2.add("button", undefined, "Cancelar", {name:"cancel"})
	
	return myWindow.show();
}

// Retorna a quantidade de partes do texto que tem um determinado estilo
function numParts(textPreference){
	clearSettings();
	if(typeof textPreference == "object"){
		app.findTextPreferences.position = Position.superscript;
	} else{
		app.findTextPreferences.fontStyle = textPreference;
	}
	return app.activeDocument.findText().length;
}

// Aplica um estilo de caractere a todos os trechos com uma determinado preferencia de texto
function applyStyle(textPreference, charStyle){
	clearSettings();
	if(typeof textPreference == "object"){
		app.findTextPreferences.position = Position.superscript;
	} else{
		app.findTextPreferences.fontStyle = textPreference;
	}

	partsList = app.activeDocument.findText();

	for (j = 0; j < partsList.length; j++){
		partsList[j].appliedCharacterStyle = charStyle;
	}	
}

// Checa se o botão "procurar estilos" foi apertado.
// Depois, passa a quantidade de trechos encontrados para cada preferencia de texto para a segunda janela.
// Se o botão "aplicar estilos" for apertado executa o for que chama a função de aplicar estilos 
if(dialog1() == 1){ 
	if(dialog2(numParts(textPreferences[0]), numParts(textPreferences[1]), numParts(textPreferences[2]), numParts(textPreferences[3])) == 1){ 																																		
		for(i = 0; i < charStyleNames.length; i++){
			applyStyle(textPreferences[i], charStyleNames[i])
		}
	}
}