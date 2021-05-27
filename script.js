let fl = "";
let n = 100;
let s = 0;
let r = 0;
let y = n;
const ps = new PerfectScrollbar("#cells");
for (let j = 0; j <= n; j++) {
    if (s > 0) {
        fl = String.fromCharCode(65 + s - 1);
    }
    let x = ""
    if (j == 25) {
        x = fl + 'Z';
        n -= 26;
        j--;
        s++;
    }
    else {
        x = fl + String.fromCharCode(65 + j);
    }
    $("#columns").append(`<div class="column-name">${x}</div>`);
}
for (let j = 1; j <= y; j++) {
    $("#rows").append(`<div class="row-name">${j}</div>`);
}

let cellData={
    "Sheet1":{}
};

let selectedSheet="Sheet1";
let totalSheets=1;
let lastlyAddedSheet=1;
let saved=true;

let defaultProperties={
    "font-family":"Noto Sans",
    "font-size": 14,
    "text":"",
    "bold":false,
    "italic":false,
    "underlined":false,
    "alignment":"left",
    "color":"#444",
    "bgcolor":"#fff"
}

for (let i = 0; i < y; i++) {
    let row = $(`<div class="cell-row"></div>`);
    for (let j = 0; j < y; j++) {
        row.append(`<div id="row-${i}-col-${j}" class="input-cell" contenteditable="false"></div>`);
    }
    $("#cells").append(row);
}

$("#cells").scroll(function (e) {
    $("#rows").scrollTop(this.scrollTop)
    $("#columns").scrollLeft(this.scrollLeft);
});

$(".input-cell").dblclick(function (e) {
    $(".input-cell.selected").removeClass("selected top bottom right left");
    $(this).addClass("selected");
    $(this).attr("contenteditable", "true");
    $(this).focus();
});
$(".input-cell").blur(function (e) {
    $(this).attr("contenteditable", "false");
    updateCellData("text",$(this).text());
});

function getRowCol(e) {
    let id = $(e).attr("id");
    let idArray = id.split("-");
    let rowID = parseInt(idArray[1]);
    let colID = parseInt(idArray[3]);
    return [rowID, colID];
}

function getTopLeftRightBottom(rowID, colID) {
    let topCell = $(`#row-${rowID - 1}-col-${colID}`);
    let leftCell = $(`#row-${rowID}-col-${colID - 1}`);
    let rightCell = $(`#row-${rowID}-col-${colID + 1}`);
    let bottomCell = $(`#row-${rowID + 1}-col-${colID}`);
    return [topCell, rightCell, leftCell, bottomCell];
}

$(".input-cell").click(function (e) {
    let [rowID, colID] = getRowCol(this);
    let [topCell, rightCell, leftCell, bottomCell] = getTopLeftRightBottom(rowID, colID);
    if ($(this).hasClass("selected") && e.ctrlKey) {
        unselectCell(this, topCell, rightCell, leftCell, bottomCell);
    }
    else {
        selectCell(this, e, topCell, bottomCell, rightCell, leftCell);
    }
});
function unselectCell(e, topCell, rightCell, leftCell, bottomCell) {
    if ($(e).attr("contenteditable") == "false") {

        if ($(e).hasClass("top")) {
            topCell.removeClass("bottom");
        }
        if ($(e).hasClass("bottom")) {
            bottomCell.removeClass("top");
        }
        if ($(e).hasClass("right")) {
            rightCell.removeClass("left");
        }
        if ($(e).hasClass("left")) {
            leftCell.removeClass("right");
        }
        $(e).removeClass("selected bottom left");
    }

}
function selectCell(ele, e, topCell, bottomCell, rightCell, leftCell) {

    if (e.ctrlKey) {
        //bottom cell Selected or not
        let bottomSelected;
        if (bottomCell) {
            bottomSelected = bottomCell.hasClass("selected");
        }
        //right cell Selected or not
        let rightSelected;
        if (rightCell) {

            rightSelected = rightCell.hasClass("selected");
        }
        //left cell Selected or not
        let leftSelected;
        if (leftCell) {
            leftSelected = leftCell.hasClass("selected");
        }
        //top cell Selected or not
        let topSelected;
        if (topCell) {
            topSelected = topCell.hasClass("selected");
        }

        if (topSelected) {
            $(ele).addClass("top");
            topCell.addClass("bottom");
        }

        if (bottomSelected) {
            $(ele).addClass("bottom");
            bottomCell.addClass("top");
        }
        if (rightSelected) {
            $(ele).addClass("right");
            rightCell.addClass("left");
        }
        if (leftSelected) {
            $(ele).addClass("left");
            leftCell.addClass("right");
        }

    }
    else {
        $(".input-cell.selected").removeClass("selected top bottom right left");
    }
    $(ele).addClass("selected");
    changeHeader(getRowCol(ele));
    
}
function changeHeader([row,col]){
    let data;
    if(cellData[selectedSheet][row]&&cellData[selectedSheet][row][col]){
        data=cellData[selectedSheet][row][col];
    }
    else{
        data=defaultProperties;
    }
    $(".alignment.selected").removeClass("selected");
    $(`.alignment[data-type=${data.alignment}]`).addClass("selected");
    if(data.bold){
        $("#bold").addClass("selected");
    }
    else{
        $("#bold").removeClass("selected");
    }
    if(data.italic){
        $("#italic").addClass("selected");
    }
    else{
        $("#italic").removeClass("selected");
    }
    if(data.underline){
        $("#underline").addClass("selected");
    }
    else{
        $("#underline").removeClass("selected");
    }
    $("#fill-color").css("border-bottom",`4px solid ${data.bgcolor}`);
    $("#text-color").css("border-bottom",`4px solid ${data.color}`);
    $("#font-family").val(data["font-family"]);
    $("#font-family").css("font-family",`${data["font-family"]}`);
    $("#font-size").val(data["font-size"]);
    
}
let startCellSelected = false;
let startCell = {};
let endCell = {};
let scrollXRstarted=false;
let scrollXLstarted=false;

$(".input-cell").mousemove(function (e) {
    e.preventDefault();
    if (e.buttons == "1") {
        if(e.pageX>$(window).width()-70 && !scrollXRstarted){
            scrollXR(e);
        }
        if(e.pageX<100 && !scrollXLstarted ){
            scrollXL(e);
        }
        if (!startCellSelected) {
            let [row, col] = getRowCol(this);
            startCell = { "rowID": row, "colID": col };
            startCellSelected = true;
            selectAllBetweenMultipleCells(startCell, startCell);
        }

    }
    else {
        startCellSelected = false;
    }
})
let count=0;
$(".input-cell").mouseenter(function (e) {
    if (e.buttons == 1) {
        if(e.pageX<$(window).width()-70 && scrollXRstarted){
            clearInterval(scrollXRinterval);
            scrollXRstarted=false;
        }
        if(e.pageX>100 && scrollXLstarted){
            clearInterval(scrollXLinterval);
            scrollXLstarted=false;
        }
        let [row, col] = getRowCol(this);
        endCell = { "rowID": row, "colID": col };
        selectAllBetweenMultipleCells(startCell, endCell);
    }
})

function selectAllBetweenMultipleCells(start, end) {
    $(".input-cell").removeClass("selected top bottom left right");
    for (let i = Math.min(start.rowID, end.rowID); i <= Math.max(start.rowID, end.rowID); i++) {
        for (let j = Math.min(start.colID, end.colID); j <= Math.max(start.colID, end.colID); j++) {
            let [topCell, rightCell, leftCell, bottomCell] = getTopLeftRightBottom(i, j);
            selectCell($(`#row-${i}-col-${j}`)[0], { "ctrlKey": true }, topCell, bottomCell, rightCell, leftCell);
        }
    }
}

let scrollXRinterval;
let scrollXLinterval;

function scrollXR(e){
    scrollXRstarted=true;
    scrollXRinterval=setInterval(()=>{
        $("#cells").scrollLeft($("#cells").scrollLeft()+100);
        
    },100);
}

function scrollXL(e){
    scrollXLstarted=true;
    scrollXLinterval=setInterval(()=>{
        $("#cells").scrollLeft($("#cells").scrollLeft()-100);
    },100);
}

$(".data-container").mouseup(function(e){
        clearInterval(scrollXRinterval);
        clearInterval(scrollXLinterval);
        scrollXLstarted=false;
        scrollXRstarted=false;
})

$(".alignment").click(function(e){
    let alignment=$(this).attr("data-type");
    $(".alignment.selected").removeClass("selected");
    $(this).addClass("selected");
    $(".input-cell.selected").css("text-align",alignment); 
    updateCellData("alignment",alignment);
})

$("#bold").click(function(e){
    setStyle(this,"font-weight","bold");

});


$("#italic").click(function(e){
    setStyle(this,"font-style","italic");

});

$("#underline").click(function(e){
   setStyle(this,"text-decoration","underline");
});

function setStyle(ele,key,value){
    
        if($(ele).hasClass("selected")){
            $(ele).removeClass("selected");
            $(".input-cell.selected").css(key,"");
            // $(".input-cell.selected").each(function(index,data){
            //     let [row,col]=getRowCol(data);
            //     cellData[row][col][value]=false;
            // });
            updateCellData(value,false);
        }
        else{
            $(ele).addClass("selected");
            $(".input-cell.selected").css(key,value);
            // $(".input-cell.selected").each(function(index,data){
            //     let [row,col]=getRowCol(data);
            //     cellData[row][col][value]=true;
            // });
            updateCellData(value,true);
        }   
}

$(".pick-color").colorPick({
    // 'initialColor': '#3498db',
    'allowRecent': true,
    'recentMax': 5,
    'allowCustomColor': true,
    'palette': ["#1abc9c", "#16a085", "#2ecc71", "#27ae60", "#3498db", "#2980b9", "#9b59b6", "#8e44ad", "#34495e", "#2c3e50", "#f1c40f", "#f39c12", "#e67e22", "#d35400", "#e74c3c", "#c0392b", "#ecf0f1", "#bdc3c7", "#95a5a6", "#7f8c8d"],
    'onColorSelected': function(e) {
        if(this.color!="ABCD"){
            if($(this.element.children()[1]).attr("id")=="fill-color"){
                let clr=this.color;
                $(".input-cell.selected").css( "background-color", this.color);
                $("#fill-color").css("border-bottom",`4px solid ${this.color}`);
                // $(".input-cell.selected").each(function(index,data){
                //     let [row,col]=getRowCol(data);
                //     cellData[row][col].bgcolor=clr;
                // })
                updateCellData("bgcolor",clr);
            }
        }
        if($(this.element.children()[1]).attr("id")=="text-color"){
            let clr=this.color;
            $(".input-cell.selected").css( "color", this.color);
            $("#text-color").css("border-bottom",`4px solid ${this.color}`)
            // $(".input-cell.selected").each(function(index,data){
            //     let [row,col]=getRowCol(data);
            //     cellData[row][col].color=clr;
            // })
            updateCellData("color",clr);
        }
    }
  });

$("#fill-color").click(function(e){
    setTimeout(()=>{
        $(this).parent().click();
    },10) 
});
$("#text-color").click(function(e){
    setTimeout(()=>{
        $(this).parent().click();
    },10) 
})

$(".font-selector").change(function(e){
    let val=$(this).val();
    console.log(val);
    let key=$(this).attr("id");
    if(key=="font-family"){
        $("#font-family").css(key,val);
    }
    if(!isNaN(val)){
        val=parseInt(val);
    }
        $(".input-cell.selected").css(key,val);
        $(".input-cell.selected").each(function(index,data){
            let [row,col]=getRowCol(data);
            cellData[row][col][key]=val;
        })
})

function updateCellData(property,value){
    let currCellData=JSON.stringify(cellData);
    if(value!=defaultProperties[property]){
        $(".input-cell.selected").each(function(index,data){
            let [row,col]=getRowCol(data);
            if(cellData[selectedSheet][row]==undefined){
                cellData[selectedSheet][row]={};
                cellData[selectedSheet][row][col]={...defaultProperties};
                cellData[selectedSheet][row][col][property]=value;
            }
            else{
                if( cellData[selectedSheet][row][col]==undefined){
                    cellData[selectedSheet][row][col]={...defaultProperties};
                    cellData[selectedSheet][row][col][property]=value;
                }
                else{
                    cellData[selectedSheet][row][col][property]=value;
                }
            }
        });
    }
    else{
        $(".input-cell.selected").each(function(index,data){
            let [row,col]=getRowCol(data);
        if(cellData[selectedSheet][row]&&cellData[selectedSheet][row][col]){
            cellData[selectedSheet][row][col].alignment=alignment;
            if(JSON.stringify(cellData[selectedSheet][row][col])==JSON.stringify(defaultProperties)){
                delete cellData[selectedSheet][row][col];
                if(Object.keys(cellData[selectedSheet][row]).length==0){
                    delete cellData[selectedSheet][row];
                }
            }
        }});
    }
    if(saved&&currCellData!=JSON.stringify(cellData)){
        saved=false;
    }
}
$(".container").click(function(e){
    $(".sheet-options-modal").remove();
})
function addSheetEvents(){
    $(".sheet-tab.selected").on("contextmenu",function(e){
        e.preventDefault();
        selectSheet(this);
        $(".sheet-options-modal").remove();
        let modal=$(`<div class="sheet-options-modal">
                        <div class="options sheet-rename">Rename</div>
                        <div class="options sheet-delete">Delete</div>
                    </div>`);
        modal.css({"left":e.pageX});
        $(".container").append(modal);
        let renameModal = $(`<div class="sheet-modal-parent">
                            <div class="sheet-modal-rename">
                                <span class="sheet-modal-title">Rename Sheet</span>
                                <div class="sheet-modal-input-container">
                                    <span class="sheet-modal-input-title">Rename ${selectedSheet} to:</span>
                                    <input class="sheet-modal-input" type="text"/>
                                </div>
                                <div class="sheet-modal-confirmation">
                                    <div class="button ok-button">Done</div>
                                    <div class="button cancel-button">Cancel</div>
                                </div>
                            </div>
                    </div>`);
        $(".sheet-rename").click(function (e) {
            $(".container").append(renameModal);
            $(".cancel-button").click(function(e){
                $(".sheet-modal-parent").remove();
            });
            $(".ok-button").click(function(e){
                renameSheet();
            });
            $(".sheet-modal-input").keypress(function(e){
                if(e.key=="Enter"){
                    renameSheet();
                }
            })
        });
        let deleteModal=$(`<div class="sheet-modal-parent">
            
        <div class="sheet-modal-delete">
            <span class="sheet-modal-title">${$(this).text()}</span>
            <div class="sheet-modal-detail">
                <span class="sheet-modal-detail-title">Are you sure?</span>    
            </div>
            <div class="sheet-modal-confirmation">
                <div class="button ok-button"> 
                <div class="material-icons">delete</div>
                Delete</div>
                <div class="button cancel-button">Cancel</div>
            </div>
        </div>
    </div>`);
        $(".sheet-delete").click(function(e){
            if(totalSheets>1){
                $(".container").append(deleteModal);
                $(".ok-button").click(function(e){
                    deleteSheet();
                });
                $(".cancel-button").click(function(e){
                    $(".sheet-modal-parent").remove();
                });
            }
            else{
                alert("not possible")
            }
        })
            
        

    });
    $(".sheet-tab.selected").click(function(e){
        
        selectSheet(this);
    });  
}
addSheetEvents();

$(".add-sheet").click(function(e){
    addSheet();    
})
function addSheet(){
    saved=false;
    lastlyAddedSheet++;  totalSheets++;
    $(".sheet-tab.selected").removeClass("selected");
    $(".sheet-tab-container").append(`<div class="sheet-tab selected">Sheet${lastlyAddedSheet}</div>`);
    cellData[`Sheet${lastlyAddedSheet}`]={};
    selectSheet();
    addSheetEvents();    
    $(".sheet-tab.selected")[0].scrollIntoView();
}
function selectSheet(ele){
    if (ele && !$(ele).hasClass("selected")) {
        $(".sheet-tab.selected").removeClass("selected");
        $(ele).addClass("selected");
    }
    emptyPreviousSheet();
    selectedSheet=$(".sheet-tab.selected").text();
    loadCurrentSheet();
}

function emptyPreviousSheet(){
    let data=cellData[selectedSheet];
    let rowKeys=Object.keys(data);
    // console.log(rowKeys);
    for(let i of rowKeys){
        let row=parseInt(i);
        let colKeys=Object.keys(data[i]);
        console.log(colKeys);
        for(let j of colKeys){
            let col=parseInt(j);
            let cellKeys=$(`#row-${row}-col-${col}`);
            cellKeys.text("");
            cellKeys.css({
                "font-family":"Noto Sans",
                "font-size": 14,
                "font-weight":"",
                "font-style":"",
                "font-decoration":"",
                "text-align":"left",
                "color":"#444",
                "background-color":"#fff"
            })
        }
    }
}

function loadCurrentSheet(){
    let data=cellData[selectedSheet];
    $("#row-0-col-0").click();
    let rowKeys=Object.keys(data);
    for(let i of rowKeys){
        let row=parseInt(i);
        let colKeys=Object.keys(data[i]);
        for(let j of colKeys){
            let col=parseInt(j);
            let cellKeys=$(`#row-${row}-col-${col}`);
            let cellD=data[row][col];
            cellKeys.text(cellD.text);
            cellKeys.css({
                "font-family":cellD["font-family"],
                "font-size": cellD["font-size"],
                "font-weight":cellD["bold"]?"bold":"",
                "font-style":cellD["italic"]?"italic":"",
                "font-decoration":cellD["underlined"]?"underline":"",
                "text-align":cellD["alignment"],
                "color":cellD["color"],
                "background-color":cellD["bgcolor"]
            })
        }
    }
}

function renameSheet(){
let newSheetName=$(".sheet-modal-input").val();
if(newSheetName && !Object.keys(cellData).includes(newSheetName)){
    saved=false;
    let newCellData = {};
        for (let i of Object.keys(cellData)) {
            if (i == selectedSheet) {
                newCellData[newSheetName] = cellData[selectedSheet];
            } else {
                newCellData[i] = cellData[i];
            }
        }

        cellData = newCellData;
    selectedSheet=newSheetName;
    $(".sheet-tab.selected").text(selectedSheet);
    $(".sheet-modal-parent").remove();
}
else{
    $(".sheet-modal-input-container").append(
        `<div class="rename-error">Sheet Name is not valid or doesn't exists!</div>`
    )
}
}

function deleteSheet() {
    $(".sheet-modal-parent").remove();
    let sheetIndex = Object.keys(cellData).indexOf(selectedSheet);
    let currSelectedSheet = $(".sheet-tab.selected");
    if (sheetIndex == 0) {
        selectSheet(currSelectedSheet.next()[0]);
    } else {
        selectSheet(currSelectedSheet.prev()[0]);
    }
    delete cellData[currSelectedSheet.text()];
    currSelectedSheet.remove();
    totalSheets--;
}

$(".left-scroller").click(function(e){
    let sheetIndex=Object.keys(cellData).indexOf(selectedSheet);
    let currSheet=$(".sheet-tab.selected");
    if(sheetIndex!=0){
        selectSheet(currSheet.prev()[0]);
    }
    $(".sheet-tab.selected")[0].scrollIntoView();
});

$(".right-scroller").click(function(e){
    let sheetIndex=Object.keys(cellData).indexOf(selectedSheet);
    let currSheet=$(".sheet-tab.selected");
    if(sheetIndex!=totalSheets-1){
        selectSheet(currSheet.next()[0]);
    }
    $(".sheet-tab.selected")[0].scrollIntoView();

});




$(".menu-bar-item").click(function(e){
    if($(this).text()=="File"){

        let fileModal=$(`<div class="file-modal">
                            <div class="first-half">
                                <div class="close">
                                    <div class="material-icons back-icon">arrow_circle_down</div>
                                Close</div> 
                                <div class="new-file">
                                    <div class="material-icons new-icon">note_add</div>
                                New</div>
                                <div class="open-file">
                                    <span class="material-icons open-icon">folder_open</span>
                                Open</div>
                                <div class="save-file">
                                    <div class="material-icons save-icon">save</div>
                                Save</div>
                            </div>
                            <div class="second-half"></div>
                            <div class="pop-file-modal"></div>

                            </div>`);

        $(".container").append(fileModal);
        $(".file-modal").animate({width:"100vw"});
        $(".pop-file-modal,.close,.new-file,.save-file,.open-file").click(function(e){
            $(".file-modal").animate({width:"0px"},300);
            setTimeout(()=>{
                fileModal.remove();
            },290);
        });

        $(".new-file").click(function(e){
            if(saved){
                newFile();
            }
            else{
                let newModal=$(`<div class="sheet-modal-parent">
                                    <div class="sheet-modal-delete">
                                        <span class="sheet-modal-title">${$(".title").text()}</span>
                                        <div class="sheet-modal-detail">
                                            <span class="sheet-modal-detail-title">Do you want to save changes?</span>    
                                        </div>
                                        <div class="sheet-modal-confirmation">
                                            <div class="button ok-button"> 
                                            Yes</div>
                                            <div class="button cancel-button">Cancel</div>
                                        </div>
                                    </div>
                                </div>`);
                $(".container").append(newModal);
                $(".ok-button").click(function(e){
                    newModal.remove();
                    saveFile(true);
                });
                $(".cancel-button").click(function(e){
                    
                    newModal.remove();newFile();
                })
                
                
            }
        });

        $(".save-file").click(function(e){
            saveFile();
        })
    }
    $(".open-file").click(function(e){
        openFile();
    })

});

function newFile(){
    emptyPreviousSheet();
    cellData={"Sheet1":{}};
    $(".sheet-tab").remove();
    $(".sheet-tab-container").append(`<div class="sheet-tab selected">Sheet1</div>`);
    selectedSheet="Sheet1";
    $(".title").text("Excel - Book");
    addSheetEvents();
    totalSheets=1;
    lastlyAddedSheet=1;
    $(`#row-0-col-0`).click();
    saved=true;
}

function saveFile(clickedNew){
    let saveModal = $(`<div class="sheet-modal-parent">
                            <div class="sheet-modal-rename">
                                <span class="sheet-modal-title">Save File</span>
                                <div class="sheet-modal-input-container">
                                    <span class="sheet-modal-input-title">File Name:</span>
                                    <input class="sheet-modal-input" type="text" value=${$(".title").text()}/>
                                </div>
                                <div class="sheet-modal-confirmation">
                                    <div class="button ok-button">Done</div>
                                    <div class="button cancel-button">Cancel</div>
                                </div>
                            </div>
                    </div>`);
    $(".container").append(saveModal);
    $(".ok-button").click(function(e){
        $(".title").text($(".sheet-modal-input").val());
        let a=document.createElement("a");
        a.href=`data:application/json;charset=utf-8,${encodeURIComponent(JSON.stringify(cellData))}`;
        a.download=$(".sheet-modal-input").val()+".json"
        $(".container").append(a);
        a.click()
        // a.remove();
        
        saved=true;
    })
    $(".ok-button,.cancel-button").click(function(e){
        saveModal.remove();
        if(clickedNew){
            newFile();
        }
    })
}

function openFile(){
        let inputFile=$(`<input accept="application/json" type="file"/>`);
        $(".container").append(inputFile);
        inputFile.click();
        inputFile.change(function(e){
            let file=e.target.files[0];
            $(".title").text(file.name.split(".json")[0])
            let reader=new FileReader();
            reader.readAsText(file);
            reader.onload=()=>{
                emptyPreviousSheet();
                console.log(reader.result);
                cellData=JSON.parse((reader.result));                
                lastlyAddedSheet=1;
                $(".sheet-tab-container").text("");
                for(let i in cellData){
                    selectedSheet=i;
                    $(".sheet-tab.selected").removeClass("selected");
                    $(".sheet-tab-container").append(`<div class="sheet-tab selected">${i}</div>`);
                    addSheetEvents();
                }
                let sheets=Object.keys(cellData);
                selectedSheet=sheets[0];
                totalSheets=sheets.length;
                lastlyAddedSheet=totalSheets;
                inputFile.remove();
            }
        });
}
let clipboard={"startCell":[],"cellData":{}};

$("#copy").click(function(e){
    clipboard.startCell=getRowCol($(".input-cell.selected")[0]);
    $(".input-cell.selected").each(function(data,index){
        let [row,col]=getRowCol(data);
        if(cellData[selectedSheet][row]&&cellData[selectSheet][row][col]){
            if(!clipboard.cellData[row]){
                clipboard.cellData[row]={};
            }
            clipboard.cellData[row][col]={...cellData[selectedSheet][row][col]};


        }
    })
});

$("#paste").click(function(e){
    let startCell=getRowCol($(".input-cell.selected")[0]);
    let rows=Object.keys(clipboard.cellData);
    for(let i of rows){
        let cols=Object.keys(clipboard.cellData[i]);
        for(let j of cols){
            let rowDistance=parseInt(i)-parseInt(clipboard.startCell[0]);
            let colDistance=parseInt(j)-parseInt(clipboard.startCell[1]);
            // if(!cellData[selectedSheet][row])
        }
    }
})