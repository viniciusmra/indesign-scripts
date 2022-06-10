/*
    Autor: Vinicius Alves - viniciusmra@gmail.com
    Data: set/2021
    Última atualização: jun/2022
    
	Descrição:
	Aplica o estilo de paragráfo escolhido pelo usuário a todas as notas de rodapé.
	Apesar do inDesing possuir essa função nativamente, dependendo do documento,
	as notas de rodapé ficam com o estilo original em "overide"
*/

currentDoc = app.activeDocument;

var paragraphStyleNames = currentDoc.paragraphStyles.everyItem().name
var styleName = ""
var numFootnotes = currentDoc.stories.everyItem().footnotes.length;

function dialog1(){
	var myWindow = new Window ("dialog", "Fromatar notas de rodapé");

	// Texto
	myWindow.add("statictext", undefined, "Notas de rodapé encontradas no documento: " + numFootnotes).alignment = "left"

	// Grupo 1
	var group1 = myWindow.add('group {alignment:["right","fill"]}')
	group1.orientation = "row";
	group1.add ("statictext", undefined, "Estilo de parágrafo para as notas de rodapé:").alignment = 'left';
	dropList = group1.add ("dropdownlist", undefined, paragraphStyleNames)
    dropList.selection = 0;

	// Grupo 2
	var group2 = myWindow.add('group {alignment:["center","fill"]}')
	group2.orientation = "row";
	applyButton = group2.add("button", undefined, "Aplicar estilo", {name:"ok"})
	applyButton.enabled = false;
	group2.add("button", undefined, "Cancelar", {name:"cancel"})

	// Funções
	// Função chamda toda vez que algum estilo é selecionado
	// Caso o estilo selecionado seja diferente do primeiro (No Paragraph Style), habilida to botão de aplicar estilo
	dropList.onChange = function(){
		if(dropList.selection != 0){
			applyButton.enabled = true;
		} else{
			applyButton.enabled = false;
		}
	}

	if(myWindow.show() == 1){
		styleName = dropList.selection.text
		return 1;
	} else{
		return 0;
	}
}

// Aplica o estilo de parágrafo para todas as notas de rodapé
if(dialog1() == 1){
    currentDoc.stories.everyItem().footnotes.everyItem().paragraphs.everyItem().appliedParagraphStyle = styleName;
}

