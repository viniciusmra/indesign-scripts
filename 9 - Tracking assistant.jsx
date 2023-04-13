// Recebe o trecho selecionado no programa
selection = app.selection[0]

selection.tracking = 0 // Reseta o tracking do trecho selecionado
linesNumber = selection.lines.length // Recebe o número de linhas do trecho selecionado

// Esse for vai diminuindo o valor de "traking" de -1 até -100,
// se houver alguma diminuição do número de linhas do trecho
// o for é parado e a caixa de dialogo é exibida perguntando se
// deseja aplicar as modificações
for (i = -1; i >= -100; i--){
    selection.tracking = i
    if(linesNumber > selection.lines.length){
        break;
    }
}

//alert("Tracking = " + selection.tracking)
function dialogBox(){
    var myWindow = new Window("dialog", "Ajuste de tracking");
    var message = myWindow.add("statictext");
    message.text = "O maior traking encontrado foi: " + selection.tracking
    

    var message2 = myWindow.add("statictext");
    message2.text = "Deseja aplicar as mudanças?";
    message2.alignment = "left"

    var myButtonGroup = myWindow.add("group");
    myButtonGroup.orientation = "row"   
    yesButton = myButtonGroup.add("button", undefined, "Sim", {name:"OK"})
    noButton = myButtonGroup.add("button", undefined, "Não", {name:"Cancel"})


    //yesButton.onClick = function() {}
    //noButton.onClick = test;
    
    return myWindow.show();
}
if(dialogBox() !== 1){
    selection.tracking = 0;
}





