/*
alert("teste")
currentDoc = app.activeDocument;
paragraphStyleNames = currentDoc.characterStyles.everyItem().name


function dialog1(){
	var window = new Window ("dialog", "Aplicar estilo de caractere");
	window.orientation = "column";

	var group0 = window.add('group {alignment:["left","fill"]}')
	var text0 = group0.add ("statictext", undefined, "Escolhas os estilos de caracteres:");

	var group1 = window.add('group {alignment:["right","fill"]}')
	group1.orientation = "row";
	var text1 = group1.add ("statictext", undefined, "Estilo antigo:");
	var dp1 = group1.add ("dropdownlist", undefined, paragraphStyleNames);
	dp1.selection = 1;

	var group2 = window.add('group {alignment:["right","fill"]}')
	group2.orientation = "row";
	var text2 = group2.add ("statictext", undefined, "Estilo novo:");
	var dp2 = group2.add ("dropdownlist", undefined, paragraphStyleNames);
	dp2.selection = 2;

	var group4 = window.add('group {alignment:["center","fill"]}')
	group4.orientation = "row";
	yesButton = group4.add("button", undefined, "OK")
	noButton = group4.add("button", undefined, "Cancelar", {name:"Cancel"})
	bold = dp1.selection.text;
	if (window.show() == 1){
		names[0] = dp1.selection.text;
		names[1] = dp2.selection.text;
		names[2] = dp3.selection.text;
		names[3] = dp4.selection.text;
		return 1;
	} else{
		return 0;
	}
}

function dialog2(b, i, bi, s){
	var window = new Window ("dialog", "Encontrados");
	window.orientation = "column";

	window.add("statictext", undefined, "Foram encontrados: ").alignment = 'left'
	
	var group1 = window.add('group {alignment:["fill","fill"]}')
	group1.orientation = "column";
	var text1 = group1.add("statictext", undefined, "Negrito: " + b);
	text1.alignment = 'left'
	var text2 = group1.add("statictext", undefined, "Itálico: " + i);
	text2.alignment = 'left'
	var text3 = group1.add("statictext", undefined, "Negrito Itálico: " + bi);
	text3.alignment = 'left'
	var text4 = group1.add("statictext", undefined, "Sobrescrito: " + s);
	text4.alignment = 'left'
	
	var group2 = window.add('group {alignment:["center","fill"]}')
	group2.orientation = "row";
	group2.add("button", undefined, "OK")
	group2.add("button", undefined, "Cancelar", {name:"Cancel"})
	
	return window.show();
}



function celarSettings(){
	app.findTextPreferences = NothingEnum.nothing;
	app.changeTextPreferences = NothingEnum.nothing;
}
function findStyles(old){
	app.findTextPreferences.appliedParagraphStyle = old;
	return app.activeDocument.findText().length;
}

function changeStyles(old, newS){
	app.findTextPreferences.appliedParagraphStyle = old;
	textlList = app.activeDocument.findText();

	for (i = 0; i < textlList.length; i++){
		textlList[i].appliedParagraphStyle = newS;
	}
}
*/

app.findTextPreferences = NothingEnum.nothing;
app.changeTextPreferences = NothingEnum.nothing;
app.findTextPreferences.appliedParagraphStyle = "Normal";
normalList = app.activeDocument.findText();

alert(normalList.length)
var normalChanges = 0;
for (i=0; i<normalList.length; i++)
{
	normalList[i].appliedParagraphStyle = "body";
	normalChanges++;
}

alert ("Changes: " + normalChanges);
