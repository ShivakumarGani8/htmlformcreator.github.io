
//function reads input xlsx file given by user and to pass data to assignGlobal function
var excelData; // Global variable
function readExcel(){
	//getting file by id
	let inputXLSXFile=document.getElementById('formFileSm');
	readXlsxFile(inputXLSXFile.files[0]).then((data)=>{assignColumnNo(data)},()=>{window.alert("! invalid file or file not selected")});
}

var fileValue=[];// An empty array which stores all the tags as a separate element
var noOfOptions; // To read no of answers for question
var i=1; // Row element which ignores 1st line of arary
var questionNumber=1;

// Below variables acts as a column numbers where later we'll compare with j to call other required methods
var classColIndex;
var iDColIndex;
var answerTypeColIndex;
var questionColIndex;
var noOfOptionsColIndex;
var optionsColIndex;
var optionsClassColIndex;
var optionsIDColIndex;

//To assign col index to variables
function assignColumnNo(data){
    excelData=data;
    try{
    for(let colIndex=0;colIndex<excelData[0].length;colIndex++)
    switch(excelData[0][colIndex].toUpperCase()){
        case "CLASS_NAME" : classColIndex=colIndex;
        break;
        case "ID":iDColIndex=colIndex;
        break;
        case "TYPE_OF_ANSWER":answerTypeColIndex=colIndex;
        break;
        case "QUESTION":questionColIndex=colIndex;
        break;
        case "NO_OF_OPTIONS":noOfOptionsColIndex=colIndex;
        break;
        case "OPTIONS":answersColIndex=colIndex;
        break;
        case "OPTIONS_CLASS":optionsClassColIndex=colIndex;
        break;
        case "OPTIONS_ID":optionsIDColIndex=colIndex;
        break;
        default: throw new Error(window.alert("Invalid inputs found in "+(colIndex+1)+"th column "));
    }
    createForm(); // calling converForm method after assigning variables
    }catch(err){
        console.log("Invalid inputs");
    }
}

function createQuestion(k,l){
    let str="\t\t<div >\n"+"\t\t<p class='"+excelData[k][classColIndex]+"' id="+excelData[k][iDColIndex]+">"+excelData[k][l]+"</p>\n\t\t</div>\n";
    fileValue.push(str);
}

//Method creates options if the type of answer is singleAnswer 
function createMultiChoiceSingleAnswer(k,l){
    let value=0;// acts as a variable for value property & for id 
    let char=65;// used for naming options in character A-Z
    fileValue.push("\t\t<ul style='list-style-type:upper-alpha'>\n");
    while(noOfOptions>0){
        value++;
        fileValue.push("\t\t<li>\n");
        const str="\t\t\t<input class='"+excelData[k][optionsClassColIndex]+"' id='"+excelData[k][optionsIDColIndex]+"' type='radio' name='"+excelData[l][optionsClassColIndex]+"' value='"+value+"'> \n\t\t\t<lable for='"+excelData[k][optionsIDColIndex]+"' >"+excelData[k][l]+"</lable>\n";
        fileValue.push(str);
        fileValue.push("\t\t</li>\n");
        noOfOptions--;
        k++;
        char++;
    }
    fileValue.push("\t\t</ul>\n");
    i=k-1;
}

function createMultiChoiceMultiAnswers(k,l){
    let value=0;// acts as a variable for value property & for id 
    let char=65;// used for naming options in character A-Z
    fileValue.push("\t\t<ul style='list-style-type:upper-alpha'>\n");
    while(noOfOptions>0){
        value++;
        fileValue.push("\t\t<li>\n");
        const str="\t\t\t<input class='"+excelData[k][optionsClassColIndex]+"' id='"+excelData[k][optionsIDColIndex]+"' type='checkbox' name='"+excelData[l][optionsClassColIndex]+"' value='"+value+"'> \n\t\t\t<lable for='"+excelData[k][optionsIDColIndex]+"' >"+excelData[k][l]+"</lable>\n";
        fileValue.push(str);
        fileValue.push("\t\t</li>\n");
        noOfOptions--;
        k++;
        char++;
    }
    fileValue.push("\t\t</ul>\n");
    i=k-1;
}

function createInputType(k,l){
    fileValue.push("\t<div>\n");
    fileValue.push("\t\t<textarea class='"+excelData[k][optionsClassColIndex]+"' id='"+excelData[k][optionsIDColIndex]+"' rows='3' cols='50%' placeholder='Enter your answer'></textarea>\n");
    console.log("called")
    fileValue.push("\t</div>\n");
}

// convers given excel into form
function createForm(){
    fileValue.push("<ol>\n");
	for( ;i<excelData.length;i++){
        for(let j=0;j<excelData[0].length;j++){
            //Calling createQuestion method to create question
            if(excelData[i][j]!=undefined){
            if(j==questionColIndex){
                fileValue.push("\t<li>\n");
                createQuestion(i,j);
            }
            //Calling createMultiChoiceSingleAnswer method to create options
            try{
            if(j==answersColIndex && excelData[i][j]!=undefined ){
                noOfOptions=excelData[i][noOfOptionsColIndex];
                switch(excelData[i][answerTypeColIndex].toUpperCase()){
                    case "USER_INPUT": createInputType(i,j); 
                    break;
                    case "MULTI_CHOICE_SINGLE_ANSWER": createMultiChoiceSingleAnswer(i,j); 
                    break;
                    case "MULTI_CHOICE_MULTI_ANSWER": createMultiChoiceMultiAnswers(i,j);
                    break;
                    default:console.log("invalid answer type");    
                }
                fileValue.push("\t</li>\n");
                questionNumber++;      
            }
        }
        catch(err){
            fileValue=[];
            throw new Error(window.alert("Invalid inputs found in "+questionNumber+"th question "));
        }
        }     
        }
    }
    fileValue.push("</ol>");
    document.getElementById('convertForm').value=true;
    window.alert("File created successfully!")
}

//creates a file and downloads in local drive
const downloadFile = () => {
    if(document.getElementById('convertForm').value=="false")
    window.alert('Please click on "Create Form button"')
    else{
    const link = document.createElement("a");
    const content = fileValue.join(' ');
    const file = new Blob([content], { type: 'text/plain' });
    link.href = URL.createObjectURL(file);
    link.download = "sample.txt";
    link.d
    link.click();
    URL.revokeObjectURL(link.href);
    }
 };

//To view form on webpage
function viewForm(){
    if(document.getElementById('convertForm').value=="false")
    window.alert('Please click on "Create Form button"')
    else{
        document.getElementById("myForm").innerHTML=fileValue.join(' ');
    }
    
}
