/*
    Script escrito em setembro de 2021 por Vinicius Alves - viniciusmra@gmail.com

    O presente código cria marcadores chamados "Section Markers" nas págias em que ele encontra
    os estilos de paragrafo selecionados, dando a opção de remover páginas especificas antes
    de executar a crianção dos marcadores.

*/
currentDoc = app.activeDocument;
var selectedParagraphID = [];
var sectionList = currentDoc.sections;

paragraphStylesNames = currentDoc.paragraphStyles.everyItem().name
paragraphStylesID = currentDoc.paragraphStyles.everyItem().id;
for(i = 0; i < currentDoc.paragraphStyleGroups.length; i++){
    for(j = 0; j < currentDoc.paragraphStyleGroups.item(i).paragraphStyles.length; j++){
        paragraphStylesNames.push("("+ currentDoc.paragraphStyleGroups.item(i).name +") "+ currentDoc.paragraphStyleGroups.item(i).paragraphStyles.item(j).name)
        paragraphStylesID.push(currentDoc.paragraphStyleGroups.item(i).paragraphStyles.item(j).id)
    }
}

//#targetengine'foo';
var myWindow = new Window ("dialog", "Gerador de Marcadores de seção");

myWindow.orientation = "row";

var group1 = myWindow.add('group {alignment:["fill","fill"]}')
group1.orientation = "column";
var text1 = group1.add ("statictext", undefined, "Escolha o estilo de paragráfo:");
text1.alignment = "left"
var group1_1 = group1.add('group')
group1_1.orientation = "row";
var myDropdown = group1_1.add ("dropdownlist", undefined, paragraphStylesNames);
myDropdown.selection = 0;
var addButton = group1_1.add ('button {text: "Adicionar >>"}');

var group1_2 = group1.add('group {alignment:["left","fill"]}');
group1_2.orientation = "column";
var checkBox = group1_2.add ("checkbox", undefined, "Apagar marcações existentes");
checkBox.value = true;

var group1_3 = group1.add('group');
group1_3.orientation = "row";
group1_3.add ("button", undefined, "OK");
group1_3.add ("button", undefined, "Cancel");

var group2 = myWindow.add('group')
group2.orientation = "column";
var text2 = group2.add ("statictext", undefined, "Estilos de paragráfo selecionados:");
text2.alignment = "left"
var myList = group2.add ("listbox", undefined,);
myList.preferredSize.width = 240;
myList.preferredSize.height = 100;

addButton.onClick = function () {
    myList.add('item',myDropdown.selection.text);
    selectedParagraphID.push(paragraphStylesID[myDropdown.selection.index])
}
if(myWindow.show() == 1){
    if(selectedParagraphID.length > 0){
        var markerTittles = [];
        var markerPages = [];
        var markerTextList = [];
        for(i = 0; i < selectedParagraphID.length; i++){
            app.findTextPreferences = NothingEnum.nothing;
            app.changeTextPreferences = NothingEnum.nothing;
            app.findTextPreferences.appliedParagraphStyle = currentDoc.paragraphStyles.itemByID(selectedParagraphID[i]);
            markerList = currentDoc.findText();
            for(j = 0; j <markerList.length; j++){
                markerTittles.push(markerList[j].contents)
                markerPages.push(markerList[j].parentTextFrames[0].parentPage)
                markerTextList.push(markerList[j].parentTextFrames[0].parentPage.name + " - (" + currentDoc.paragraphStyles.itemByID(selectedParagraphID[i]).name + ") " + markerList[j].contents)
            }
        }
        
        var myWindow2 = new Window ("dialog", "Gerador de Marcadores de seção");
        var text2 = myWindow2.add("statictext", undefined, "Marcadores encontrados:");
        text2.alignment = "left"
        var myList2 = myWindow2.add ("listbox", undefined, markerTextList, {multiselect:true});
        var buttonsGroup = myWindow2.add('group {alignment:["left","fill"]}');
        var removeButton = buttonsGroup.add ('button {text: "Remover"}');
        var testButton = buttonsGroup.add ('button {text: "OK"}');
        var testButton = buttonsGroup.add ('button {text: "Cancel"}');
        removeButton.onClick = function () {
            var sel = myList2.selection[0].index;
            for (var i = 0; i < myList2.selection.length; i++){
                myList2.selection[i].text = "[removido]";
                alert(markerTittles[myList2.selection[i].index] + " - " + markerPages[myList2.selection[i].index].name);
                markerTittles.splice(myList2.selection[i].index,1);
                markerPages.splice(myList2.selection[i].index,1);
            }
        }
        if(myWindow2.show() == 1){
            if(checkBox.value == 1){
                clearMarkers();
            }
            for (j = 0; j < markerTittles.length; j++){
                try{
                        sectionList.add(markerPages[j], {sectionPrefix:"", includeSectionPrefix:false,marker:markerTittles[j]});
                }
                catch(e){}
            }
        }
        else{
            exit();
        }
    }else if(checkBox.value == 1){
        clearMarkers();
    }

} else{
    exit();
}
function clearMarkers(){
        for(i = (sectionList.length-1); i > 0; i--){
            sectionList.item(i).remove();
        }
    }
