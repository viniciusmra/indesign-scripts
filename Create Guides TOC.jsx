var currentDoc = app.activeDocument;
var txf = currentDoc.textFrames;
var selectedParagraphID = [];
var guidesLocations = []
var guidesPages = []
paragraphStylesNames = currentDoc.paragraphStyles.everyItem().name
paragraphStylesID = currentDoc.paragraphStyles.everyItem().id;
for(i = 0; i < currentDoc.paragraphStyleGroups.length; i++){
    for(j = 0; j < currentDoc.paragraphStyleGroups.item(i).paragraphStyles.length; j++){
        paragraphStylesNames.push("("+ currentDoc.paragraphStyleGroups.item(i).name +") "+ currentDoc.paragraphStyleGroups.item(i).paragraphStyles.item(j).name)
        paragraphStylesID.push(currentDoc.paragraphStyleGroups.item(i).paragraphStyles.item(j).id)
    }
}

var myWindow = new Window ("dialog", "Gerar guias");

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
        if(checkBox.value == 1){
            if(currentDoc.guides.length > 0){
                currentDoc.guides.everyItem().remove();
            }
        }
        for(i = 0; i < selectedParagraphID.length; i++){
            app.findTextPreferences = NothingEnum.nothing;
            app.changeTextPreferences = NothingEnum.nothing;
            app.findTextPreferences.appliedParagraphStyle = currentDoc.paragraphStyles.itemByID(selectedParagraphID[i]);
            markerList = currentDoc.findText();
            for(j = 0; j < markerList.length; j++){
                // Essa parte basicamente faz uma busca pelos parágrafos que o usuário selecionou (espersa-se que ele selecione os parágrafos dos titulos da TOC)
                // então são geradas duas listas: uma com a localização (em pt) do começo desses titulos
                // e outra com as páginas desses títulos 
                guidesLocations.push((markerList[j].paragraphs[0].characters[0].baseline - getHeight(markerList[j].paragraphs[0].characters[0])))
                guidesPages.push(markerList[j].paragraphs[0].parentTextFrames[0].parentPage.name)
            }
        }
    }
} else{
    exit();
}
//asd = guidesPage.sort(function (a, b) {  return a - b;  });
alert_scroll("Object properties", guidesPages);
alert_scroll("bvlalssdf", guidesLocations)
createGuides(guidesLocations,guidesPages)

function getHeight(ch){

    app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
    var frame = app.documents[0].pages[0].textFrames.add ({geometricBounds: [0, 0, 100, 100], contents: 'x'});
    frame.paragraphs[0].pointSize = ch.pointSize;
    //frame.textFramePreferences.firstBaselineOffset = FirstBaseline.X_HEIGHT;
    //o.xHeight = frame.characters[0].baseline;
    frame.textFramePreferences.firstBaselineOffset = FirstBaseline.CAP_HEIGHT;
    //o.capHeight = frame.characters[0].baseline;
    capHeight = frame.characters[0].baseline;
    frame.remove();
    return capHeight;
}
/*
function createGuides(prg){
    with (prg.parentTextFrames[0].parentPage.guides.add (app.activeDocument.activeLayer)) {
        orientation = HorizontalOrVertical.HORIZONTAL;
        location = prg.characters[0].baseline - getHeight(prg.characters[0]);
        //alert(location*0.3527777778)
    }
}
*/




function alert_scroll (title, input){
    if (input instanceof Array)
        input = input.join ("\r");
    var w = new Window ("dialog", title);
    var list = w.add ("edittext", undefined, input, {multiline: true, scrolling: true});
    list.maximumSize.height = w.maximumSize.height/2;
    list.minimumSize.width = 250;
    w.add ("button", undefined, "Close", {name: "ok"});
    w.show ();
 }

 // display properties

function createGuides(locations, pages){
    // Retira as páginas duplicadas na lista de páginas
    var uniquePages = checkUnique(pages)
    var sortLocations = []

    for(i = 0; i < uniquePages.length; i++){
        var currentLocations = []
        for(j = 0; j < pages.length; j++){
            if(uniquePages[i] == pages[j]){
                currentLocations.push(locations[j])
            }
        }
        currentLocations.sort(function (a, b) {  return a - b;  });
        for(j = 0; j < currentLocations.length; j++){
            sortLocations.push(currentLocations[j])
        }
    }

    //addGuides(sortLocations)
    addGuides(locations, pages)

}


function checkUnique(arr){
    var unique = []

    for(i = 0; i < arr.length; i++){
        var find = false
        for(j = 0; j < unique.length; j++){
            if(arr[i] == unique[j]){
                find = true
                break
            }
        }
        if(!find){
            unique.push(arr[i])
        }
    }
    unique.sort(function (a, b) {  return a - b;  });
    return unique;
}

function addGuides(loc, pages){
    selectedParagraphs = currentDoc.selection[0].paragraphs;
    selectedParagraphsPage = currentDoc.selection[0].paragraphs[0].parentTextFrames[0].parentPage.name

    // Pegar apenas as localizações que estão na presente página
    newLoc = []
    for( i = 0; i < loc.length; i++){
        if(pages[i] == selectedParagraphsPage){
            newLoc.push(loc[i])
        }
    }

    alert(newLoc)

    //alert(currentDoc.selection[0].characters.length)
    //alert_scroll("Object properties", currentDoc.selection[0].characters[0].contents.reflect.properties.sort());
    wordHeight = getHeight(currentDoc.selection[0].characters[0])
    
    // Zera todos os espaçamentos
    for(i = 0; i < selectedParagraphs.length; i++){
        selectedParagraphs[i].sameParaStyleSpacing = 0
    }

    for(i = 1; i < selectedParagraphs.length; i++){
        var a = newLoc[i-1]
        var b = (selectedParagraphs[i].characters[0].baseline - wordHeight)

        if(a - b < 0){
            alert("NUMERO NEGATIVO: " + a + " ---- " + b)
            continue
        }

        selectedParagraphs[i-1].sameParaStyleSpacing = a-b
    }

    // O que essa parte faz é pegar quais paragrafos estão selecionados (no caso os numeros da TOC)
    // Verifica em que página eles estão (espera-se que numeros de so uma página estejam selecionados)
    // Verifica o tamanho e a posição desses numeros
    // Calcula qual o ajuste que deve ser feito no numero anterior, para que o numero atual fique na posição certa
    // Esse ajuste é então passado para o parametro 


    // //alert(currentDoc.selection[0].characters.length)
    // //alert_scroll("Object properties", currentDoc.selection[0].characters[0].contents.reflect.properties.sort());
    // wordHeight = getHeight(currentDoc.selection[0].characters[0])
    
    // // Zera todos os espaçamentos
    // for(i = 0; i < selectedParagraphs.length; i++){
    //     selectedParagraphs[i].sameParaStyleSpacing = 0
    // }

    // for(i = 1; i < selectedParagraphs.length; i++){
    //     //if(i != 8){
    //         if(loc[i-1] - (selectedParagraphs[i].characters[0].baseline - wordHeight) < 0){
    //             alert("NUMERO NEGATIVO: " + loc[i-1] + " ---- " + (selectedParagraphs[i].characters[0].baseline - wordHeight))
    //             continue
    //         }
    //         var a = loc[i-1]
    //         var b = (selectedParagraphs[i].characters[0].baseline - wordHeight)
    //         selectedParagraphs[i-1].sameParaStyleSpacing = a-b
    //         //alert(a-b)
    //    //}
    // }
}