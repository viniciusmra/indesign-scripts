/*
    Autor: Vinicius Alves - viniciusmra@gmail.com
    Data: 2021
    Última atualização: jun/2022
    
    Descrição:  
    Este script tem a função de importar e linkar múltiplos arquivos do word (.doc e .docx) dentro do inDesign.
    Por padrão os documentos são importados, cada uma e uma página e depois linkados uns aos outros.
    No entanto, existe a opção de autoflow, que faz com que os documentos importados ocupem o número de páginas necessárias
    para que não ocorra "overset text". Essa opção pode ocasionar problemas em alguns tipos de documentos.
*/

var currentDoc = app.activeDocument;
var nPages = currentDoc.pages.length;
var pages = currentDoc.pages.everyItem().getElements();
var pageList = currentDoc.pages.everyItem().name;
var currentPage = currentDoc.pages.item(0);
var x = currentPage.marginPreferences.left;
var y = currentPage.marginPreferences.top;
var files;

//INTERFACE
var myWindow = new Window ("dialog", "Importar arquivos"); // Cria a janela da interface

// Grupo 1
var group1 = myWindow.add('group {alignment:["fill","fill"]}')
group1.orientation = "row";
group1.add("statictext", undefined, "Selecione os arquivos que deseja importar");
var selectButton = group1.add ('button {text: "Selecionar"}');

// ListBox
myWindow.add("statictext", undefined, "Arquivos selecionados:").alignment = "left";
var fileListBox = myWindow.add("listbox", undefined);
fileListBox.preferredSize.width = 500;
fileListBox.preferredSize.height = 300;

// Grupo 2
var group2 = myWindow.add('group {alignment:["center","fill"]}')
group2.orientation = "row";
var moveUpButton = group2.add ('button {text: "Mover para cima"}');
var moveDownButton = group2.add ('button {text: "Mover para baixo"}');
var delButton = group2.add ('button {text: "Remover arquivo"}');
moveDownButton.enabled = false;
moveUpButton.enabled = false;
delButton.enabled = false;

// Grupo 3
var group3 = myWindow.add('group {alignment:["left","fill"]}')
group3.orientation = "row";
group3.add("statictext", undefined, "Inserir arquivos após a página: ");
var dropdownList = group3.add ("dropdownlist", undefined, pageList);
dropdownList.selection = nPages-1;

//CheckBox
var checkBox = myWindow.add ("checkbox", undefined, "Fazer \"autoFlow\"");
checkBox.alignment = "left";

// Grupo 4
var group4 = myWindow.add('group {alignment:["center","fill"]}')
group4.orientation = "row";
var okButton = group4.add ("button", undefined, "Importar", {name: "ok"});
okButton.enabled = false;
group4.add ("button", undefined, "Cancelar", {name: "cancel"});

// Funções
// Abre a caixa de diálogo para selecionar os arquivos
selectButton.onClick = function () {
    if(files == null){
        files = File.openDialog("Escolha os arquivos", "*.doc;*.docx", true);
    } else{
        var temp = File.openDialog("Escolha os arquivos", "*.doc;*.docx", true);
        for(var i = 0; i < temp.length; i++){
            files.push(temp[i]);
        }
    }
    updateListBox(files,fileListBox);
}

// Move o item selecionado na listBox para cima
moveUpButton.onClick = function () {
    var index = fileListBox.selection.index;
    if(moveUp(files, fileListBox.selection.index)){
        updateListBox(files,fileListBox);
        fileListBox.selection = index - 1;
    }
}
// Move o item selecionado na listBox para baixo
moveDownButton.onClick = function () {
    var index = fileListBox.selection.index;
    if(moveDown(files, fileListBox.selection.index)){
        updateListBox(files,fileListBox);
        fileListBox.selection = index + 1;
    }   
}

// Remove o item selecionado da listBox
delButton.onClick = function () {
    var index = fileListBox.selection.index;
    
    if(fileListBox.selection != null){
        files.splice(index,1);
    }

    updateListBox(files,fileListBox);

    if(index == files.length){
        fileListBox.selection = files.length - 1;
    } else{
        fileListBox.selection = index;
    }
}

// Roda o script caso o botão importar tenha sido pressionado
if(myWindow.show() == 1){
    var pageNum = dropdownList.selection.index;
    if(files != null){
        if(files.length > 0){
            importFiles(files, pageNum);
            linkFrames(pageNum + 1,files.length)
        } 
    }
}

// Move o elemento no índice "index" para a posição "index - 1" caso ela exista
function moveUp(arr, index){
    if(index > 0){
        var temp = arr[index-1];
        arr[index-1] = arr[index];
        arr[index] = temp;
        return true;
    } 
    return false
}

// Move o elemento no índice "index" para a posição "index + 1" caso ela exista
function moveDown(arr, index){
    if(index < arr.length-1){
        var temp = arr[index+1];
        arr[index+1] = arr[index];
        arr[index] = temp;
        return true;
    } 
    return false
}

// Remove todos os elementos da listBox e coloca novos elementos
function updateListBox(items, list){
    list.removeAll();
    if(items.length > 0){
        for(var i = 0; i < items.length; i++){
            list.add('item',items[i].displayName);
        }
    }
    if(items.length == 0){
        okButton.enabled = false;
        moveDownButton.enabled = false;
        moveUpButton.enabled = false;
        delButton.enabled = false;  
    } else{
        okButton.enabled = true;
        moveDownButton.enabled = true;
        moveUpButton.enabled = true;
        delButton.enabled = true;  
    }
}

// Importa os arquivos e coloca cada um em uma página
function importFiles(files, pageNum){
    var currentPage = pages[pageNum]
    for(var i = 0; i < files.length; i++){
        currentDoc.pages.add(LocationOptions.AFTER, currentPage)
        currentPage = currentDoc.pages.item(pageNum + 1);
        currentPage.place(File(files[i]), [x,y], undefined, false,false)[0];
        pageNum++;

    }
}

// junta os Text Frames e faz "autoflow" caso a checkBox tenha sido marcada pelo usuário
function linkFrames(pageNum, nFiles){
    var allPages = currentDoc.pages;                        // Cria um lista com todas as páginas do documento
    var currentFrame = allPages[pageNum].textFrames[0];     // Define o Frame atual como sendo o primeiro Frame da página atual
    
    // Linka os arquivos recem importados
    for(var i = 0; i < nFiles-1; i++){
            pageNum++;
            currentFrame.nextTextFrame = allPages[pageNum].textFrames[0];
            currentFrame = allPages[pageNum].textFrames[0];
    }
    // Verifica se o usuário marcou o checkBox e executa o autoflow
    if(checkBox.value == 1){
        var currentPage = currentDoc.pages.item(pageNum);
        while(currentFrame.overflows){
            currentDoc.pages.add(LocationOptions.AFTER, currentPage);   // Cria uma página depois da página atual
            nextPage = allPages[++pageNum];                             // Define a próxima página
            nextFrame = addFrame(nextPage);                             // Adiciona um novo Frame a página, o nextFrame
            currentFrame.nextTextFrame = nextFrame;                     // Linka o Frame da página atual com o próximo frame (nextFrame)
            
            currentPage = nextPage; // Itera sobre as páginas
            currentFrame = currentPage.textFrames[0];
        }
    }
}

// Adiciona um Frame a página indicada e depois retorna esse frame
// Retirado de https://community.adobe.com/t5/indesign-discussions/script-to-autoflow-multiple-text-files/m-p/6214229/highlight/true
function addFrame(page){
    var pageMargins = page.marginPreferences;
    frame = page.textFrames.add();
    frameTop = page.bounds[0] + pageMargins.top;
    frameLeft = page.bounds[1] + pageMargins.left;
    frameBottom = page.bounds[2] - pageMargins.bottom;
    frameRight = page.bounds[3] - pageMargins.right;
    frame.geometricBounds = [frameTop, frameLeft, frameBottom, frameRight];
    return frame;
}