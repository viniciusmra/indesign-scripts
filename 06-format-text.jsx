//@include showProps.jsx
var currentDoc = app.activeDocument;
var spreads = currentDoc.spreads;

var pages = currentDoc.pages.everyItem().getElements();
var currentPageIndex = 0;
var currentTextFrameIndex = 0;
var currentParagraphIndex = 0;

try{
    //getSelection();
    //showProps(pages[currentPageIndex].textFrames[currentTextFrameIndex].paragraphs[currentParagraphIndex])
    //alert(pages[currentPageIndex].textFrames[currentTextFrameIndex].paragraphs[currentParagraphIndex].tracking)
} 
catch(e){}

// Lista de estilos de parágrafos
paragraphStylesNames = currentDoc.paragraphStyles.everyItem().name
paragraphStylesID = currentDoc.paragraphStyles.everyItem().id;
for(i = 0; i < currentDoc.paragraphStyleGroups.length; i++){
    for(j = 0; j < currentDoc.paragraphStyleGroups.item(i).paragraphStyles.length; j++){
        paragraphStylesNames.push("("+ currentDoc.paragraphStyleGroups.item(i).name +") "+ currentDoc.paragraphStyleGroups.item(i).paragraphStyles.item(j).name)
        paragraphStylesID.push(currentDoc.paragraphStyleGroups.item(i).paragraphStyles.item(j).id)
    }
}

// Necessário para o "palette" funcionar (as barras "//" servem pro vscode não apontar erro)
//@targetengine "mysession"

// Main
palette1();

function palette1(){
    var myWindow = new Window ("palette", "Formatador");
    //myWindow.spacing = 0

    myWindow.add("statictext", undefined, "Navegação");

    // Group 1
    var group1 = myWindow.add('group {alignment:["center","fill"]}')
    group1.orientation = "row";
    previousParagraphButton = group1.add("button", undefined, "<<");
    group1.add("statictext", undefined, "Parágrafo");
    nextParagraphButton = group1.add("button", undefined, ">>");

    // Group 2
    var group2 = myWindow.add('group {alignment:["center","fill"]}')
    group2.orientation = "row";
    previousPageButton = group2.add("button", undefined, "<<");
    group2.add('statictext {text: "Página", characters: 10, justify: "center"}').preferredSize.width = 58;
    nextPageButton = group2.add("button", undefined, ">>");

    var e = myWindow.add("edittext" , [0,0,100,24] , "");
	e.addEventListener("keyup" , keypress);
    decreaseTrackingButton = myWindow.add("button", undefined, "-");
    increaseTrackingButton = myWindow.add("button", undefined, "+");

    myWindow.add("statictext", undefined, "Estilos de parágrafo");

    // Group 3
    var group3 = myWindow.add('group {alignment:["center","fill"]}')
    group3.orientation = "row";
    var group3_1 = group3.add('group {alignment:["left","fill"]}')
    group3_1.orientation = "column";
    group3_1.spacing = 5
    var group3_2 = group3.add('group {alignment:["left","fill"]}')
    group3_2.orientation = "column";
    group3_2.spacing = 5
    
    var styleButtons = [];
    var shortcutEditTexts = [];
    var shortcuts = [];
    for(var i = 0; i < paragraphStylesNames.length; i++){
        if(i < (paragraphStylesNames.length)/2){
            var group3_1_1 = group3_1.add('group {alignment:["center","fill"]}')
            group3_1_1.orientation = "row";
            shortcutEditTexts.push(group3_1_1.add("edittext" , [0,0,20,24] , ""));
            styleButtons.push(group3_1_1.add("button", undefined, paragraphStylesNames[i]));
        } else{
            var group3_2_1 = group3_2.add('group {alignment:["center","fill"]}')
            group3_2_1.orientation = "row";
            styleButtons.push(group3_2_1.add("button", undefined, paragraphStylesNames[i]));
            shortcutEditTexts.push(group3_2_1.add("edittext" , [0,0,20,24] , ""));
        }
        styleButtons[i].characters = 20;
        styleButtons[i].onClick = styleButtonClick;
        shortcuts.push(null);
        shortcutEditTexts[i].onChanging = SetShortcut;
        
    }
    function styleButtonClick(){
        applyStyle(this.text);
    }

    function SetShortcut(){
        //alert(paragraphStylesNames.length)
        //alert(shortcuts.length)
        //alert(shortcutEditTexts.length)
        for(var i = 0; i < paragraphStylesNames.length; i++){
            if(shortcutEditTexts[i] === this){
                //alert(i)
                shortcuts[i] = this.text[0].toUpperCase();
                
            }else if(shortcutEditTexts[i].text[0] === this.text[0]){
                //alert("asd")
                shortcutEditTexts[i].text = "";
                shortcuts[i] = null;
            }
        }
    }

    increaseTrackingButton.onClick = tracking;
    decreaseTrackingButton.onClick = tracking;
    decreaseTrackingButton.onClick = tracking;

    nextParagraphButton.onClick = function(){
        getSelection();
        nextParagraph();
    }

    previousParagraphButton.onClick = function(){
        getSelection();
        previousParagraph();
    }

    nextPageButton.onClick = function(){
        getSelection();
        nextPage();
    }

    previousPageButton.onClick = function(){
        getSelection();
        previousPage()
    }

    function keypress(k) {
        var keyPressed = k.keyName.toUpperCase();
        e.active = false;
        getSelection();
        if(keyPressed == "DOWN"){
            nextParagraph();
        }
        if(keyPressed == "UP"){
            previousParagraph();
        }
        //alert(keyPressed)
        for(var i = 0; i < shortcuts.length; i++) {
            //alert(k.keyName.toUpperCase() + " - " + shortcuts[i]);
            if(keyPressed == shortcuts[i]){
                //alert("achei")
                applyStyle(paragraphStylesNames[i])
            }
        }

        e.active = true;
        /*
		str = [];
		if (e.altKey) str.push("Alt");
		if (e.ctrlKey) str.push("Ctrl");
		if (e.shiftKey) str.push("Shift");
		if (e.metaKey) str.push("Win");
		str.push(k.keyName);
		alert(str);*/
	}

    return myWindow.show();
}


// A função verifica se a combinação de page, textFrame e paragraph se existe no documento e retorna um código numérico
// - 0: Dentro dos intervalos
// - 1: Fora do intervalo de Pages
// - 2: Fora do intervalo de TextFrames
// - 3: Fora do intervalo de Paragraphs
function paragraphVerification(page, textFrame, paragraph) {
    // Verifica se está dentro do intervalo de páginas
    if(page >= pages.length || page < 0){
        return 1;
    }
    
    // Verifica se está dentro do intervalo de TextFrames
    if(textFrame >= pages[page].textFrames.length || textFrame < 0){
        return 2;
    }
    
    // Verifica se está dentro do intervalo de parágrafos
    if(paragraph >= pages[page].textFrames[textFrame].paragraphs.length || paragraph < 0){
        return 3;
    }
    // Retorna 0 se o parágrafo existe
    return 0;
}

function nextParagraph(){
    //var breakParagraph = false;
    //if(pages[currentPageIndex].textFrames[currentTextFrameIndex].paragraphs[currentParagraphIndex].parentTextFrames.length > 1){
    //    breakParagraph = true;
    //}
    currentParagraphIndex++;
    var verification = paragraphVerification(currentPageIndex, currentTextFrameIndex, currentParagraphIndex);
    do{
        if(verification == 1){
            return 0;
        } else if(verification == 2){
            currentTextFrameIndex = 0;
            currentPageIndex++;
        } else if(verification == 3){
            currentParagraphIndex = 0;
            currentTextFrameIndex++;
        }
        if(paragraphVerification(currentPageIndex, currentTextFrameIndex, currentParagraphIndex) == 0 && currentParagraphIndex == 0 && pages[currentPageIndex].textFrames[currentTextFrameIndex].paragraphs[currentParagraphIndex].parentTextFrames.length > 1){
            currentParagraphIndex++
        }
        verification = paragraphVerification(currentPageIndex, currentTextFrameIndex, currentParagraphIndex);
    }while(verification != 0);

    //pages[currentPageIndex].textFrames[currentTextFrameIndex].paragraphs[currentParagraphIndex].showText(); // Vai até o parágrafo no documento
    var sp = findSpread(currentPageIndex);
    sp.select();
    currentDoc.selection = pages[currentPageIndex].textFrames[currentTextFrameIndex].paragraphs[currentParagraphIndex];
}

function previousParagraph(){
    currentParagraphIndex--;
    var verification = paragraphVerification(currentPageIndex, currentTextFrameIndex, currentParagraphIndex);
    while(verification != 0){
        if(verification == 1){
            return 0;
        } else if(verification == 2){
            if(currentPageIndex > 0){
                currentPageIndex--;
                if(pages[currentPageIndex].textFrames.length > 0){
                    currentTextFrameIndex = pages[currentPageIndex].textFrames.length - 1;
                    if(pages[currentPageIndex].textFrames[currentTextFrameIndex].paragraphs.length > 0){
                        currentParagraphIndex = pages[currentPageIndex].textFrames[currentTextFrameIndex].paragraphs.length - 1;
                    }
                }
                
            } else{
                currentPageIndex = 0;
                currentTextFrameIndex = 0;
                currentParagraphIndex = 0;
                return 0;
            }
            currentTextFrameIndex = pages[currentPageIndex].textFrames.length-1;
        } else{
            currentTextFrameIndex--;
            currentParagraphIndex = pages[currentPageIndex].textFrames[currentTextFrameIndex].paragraphs.length - 1;
            
        }
        verification = paragraphVerification(currentPageIndex, currentTextFrameIndex, currentParagraphIndex);
    }

    //pages[currentPageIndex].textFrames[currentTextFrameIndex].paragraphs[currentParagraphIndex].showText(); // Vai até o parágrafo no documento
    var sp = findSpread(currentPageIndex);
    sp.select();
    currentDoc.selection = pages[currentPageIndex].textFrames[currentTextFrameIndex].paragraphs[currentParagraphIndex];
}

function nextPage(){
    currentPageIndex++;
    currentTextFrameIndex = 0;
    currentParagraphIndex = -1;
    return nextParagraph();
}
function previousPage(){
    currentParagraphIndex = 0;
    currentTextFrameIndex = 0;
    return previousParagraph();
}

function getSelection(){
    var userSelection = currentDoc.selection[0];

    // Verifica se tem algum objeto selecionado, caso não tenha encerra a função
    if(userSelection == undefined){
        currentPageIndex = 0;
        currentTextFrameIndex = 0;
        currentParagraphIndex = 0;
        return;
    }

    // Checa se a seleção é uma parte do texto usando "instanceof".
    // Quando o usuário seleciona o texto, a seleção pode ser dos tipos:
    // - InsertionPoint: O cursor no texto mas sem nada selecionado
    // - Character: Um caractere é selecionado
    // - Word: Uma palavra é selecionada
    // - Paragraph: Um parágrafo inteiro é selecionado (incluindo o caractere de quebra de linha)
    // - Text: Tudo que não se encaixar nos demais, incluindo vários parágrafos
    var isText = false
    if(userSelection instanceof InsertionPoint || userSelection instanceof Character || userSelection instanceof Word || userSelection instanceof Paragraph || userSelection instanceof Text || userSelection instanceof TextStyleRange || userSelection instanceof TextColumn){
        isText = true; 
        var currentParagraph = userSelection.paragraphs[0];         // Atribui o [Object Paragraph] à variável "currentParagraph"
        currentDoc.selection = currentParagraph;
        var currentTextFrame = userSelection.parentTextFrames[0];    // Atribui o [Object TextFrame] à variável "currentTextFrame"
        var currentPage = currentTextFrame.parentPage;               // Atribui o [Object Page] ao qual o objeto selecionado está associado à variável "currentPage"
        currentPageIndex = currentPage.name - 1;                     // Atribui o índice da página atual

    }

    else{ // Caso não seja um texto
        var currentPage = userSelection.parentPage;             // Atribui o [Object Page] ao qual o objeto selecionado está associado à variável "currentPage"
        currentPageIndex = userSelection.parentPage.name - 1;   // Atribui o índice da página atual
        
        // Caso seja uma caixa de texto
        if(userSelection instanceof TextFrame){
            var currentTextFrame = currentDoc.selection[0];
        
        } else{ // Caso não seja uma caixa de texto (pode ser uma imagem, ou forma), atribui p índice 0 e encerra a função
            currentTextFrameIndex = 0;
            currentParagraphIndex = 0; 
            return;
        }
    }

    // Busca o índice do TextFrame selecionado dentro dos TextFrames da página atual
    for(var i = 0; i < currentPage.textFrames.length; i++){
        if(currentPage.textFrames[i] === currentTextFrame){
            currentTextFrameIndex = i;
        }
    }

    // Caso seja um texto, busca o índice do parágrafo selecionado dentro dos Paragraphs do TextFrame atual
    if(isText){
        for(var i = 0; i < currentTextFrame.paragraphs.length; i++){
            if(currentTextFrame.paragraphs[i] === currentParagraph){
                currentParagraphIndex = i;
            }
        }
    }
}

function findParagraphStyleID(name){
    for(var i = 0; i < paragraphStylesNames.length; i++){
        if(paragraphStylesNames[i] === name){
            return paragraphStylesID[i];
        }
    }
}

function findSpread(index) {
    for(var i = 0; i < spreads.length; i++){
        for(var j = 0; j < spreads[i].pages.length; j++){
            if(index === spreads[i].pages[j].name-1){
                return spreads[i];
            }
        }
    }
}

function applyStyle(style){
    getSelection();
    pages[currentPageIndex].textFrames[currentTextFrameIndex].paragraphs[currentParagraphIndex].appliedParagraphStyle = currentDoc.paragraphStyles.itemByID(findParagraphStyleID(style));
    nextParagraph()
}

function tracking(){
    getSelection();
    var tracking = 0;
    if(pages[currentPageIndex].textFrames[currentTextFrameIndex].paragraphs[currentParagraphIndex].lines.length > 0){
        tracking = pages[currentPageIndex].textFrames[currentTextFrameIndex].paragraphs[currentParagraphIndex].tracking;
        pages[currentPageIndex].textFrames[currentTextFrameIndex].paragraphs[currentParagraphIndex].tracking = tracking - 1;
    }
}