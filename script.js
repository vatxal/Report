const startupScreen = document.getElementById("startupScreen");
const mainContainer = document.getElementById("mainContainer");

const continueBtn = document.getElementById("continueBtn");
const newBtn = document.getElementById("newBtn");

const rowsContainer = document.getElementById("rowsContainer");

const addBtn = document.getElementById("addBtn");
const submitBtn = document.getElementById("submitBtn");

let projectData = [];

/* ------------------------------ */
/* STORAGE */
/* ------------------------------ */

function saveProject(){

  localStorage.setItem(
    "exceljs_project",
    JSON.stringify(projectData)
  );
}

function loadProject(){

  const saved = localStorage.getItem("exceljs_project");

  if(saved){

    projectData = JSON.parse(saved);

    rowsContainer.innerHTML = "";

    const temp = [...projectData];

    projectData = [];

    temp.forEach(item => {
      createRow(item.image, item.note);
    });
  }
}

/* ------------------------------ */
/* CREATE ROW */
/* ------------------------------ */

function createRow(imageSrc = "", note = ""){

  const index = projectData.length;

  projectData.push({
    image:imageSrc,
    note:note
  });

  const card = document.createElement("div");

  card.className = "row-card";

  card.innerHTML = `

    <div class="row-layout">

      <div class="left-section">

        <img class="image-preview"
             src="${imageSrc}" />

        <input
          type="file"
          accept="image/*"
          capture="environment"
          class="file-input"
        />

      </div>

      <div class="right-section">

        <textarea
          placeholder="Write notes here..."
        >${note}</textarea>

      </div>

    </div>

  `;

  const img = card.querySelector(".image-preview");

  const fileInput = card.querySelector(".file-input");

  const textarea = card.querySelector("textarea");

  /* IMAGE */

  fileInput.addEventListener("change",(e)=>{

    const file = e.target.files[0];

    if(!file) return;

    const reader = new FileReader();

    reader.onload = function(event){

      img.src = event.target.result;

      projectData[index].image = event.target.result;

      saveProject();
    };

    reader.readAsDataURL(file);
  });

  /* NOTE */

  textarea.addEventListener("input",()=>{

    projectData[index].note = textarea.value;

    saveProject();
  });

  rowsContainer.appendChild(card);

  saveProject();
}

/* ------------------------------ */
/* NEW PROJECT */
/* ------------------------------ */

newBtn.addEventListener("click",()=>{

  localStorage.removeItem("exceljs_project");

  projectData = [];

  rowsContainer.innerHTML = "";

  createRow();

  startupScreen.classList.add("hidden");

  mainContainer.classList.remove("hidden");
});

/* ------------------------------ */
/* CONTINUE */
/* ------------------------------ */

continueBtn.addEventListener("click",()=>{

  startupScreen.classList.add("hidden");

  mainContainer.classList.remove("hidden");

  loadProject();

  if(projectData.length === 0){
    createRow();
  }
});

/* ------------------------------ */
/* ADD */
/* ------------------------------ */

addBtn.addEventListener("click",()=>{

  createRow();
});

/* ------------------------------ */
/* EXCEL EXPORT */
/* ------------------------------ */

submitBtn.addEventListener("click", async ()=>{

  const workbook = new ExcelJS.Workbook();

  const worksheet = workbook.addWorksheet("Sheet1");

  /* ------------------------------ */
  /* COLUMN WIDTHS */
  /* ------------------------------ */

  /*
    ExcelJS width units are not exact cm.
    15cm ≈ 60 width
  */

  worksheet.getColumn(1).width = 60;
  worksheet.getColumn(2).width = 60;

  /* ------------------------------ */
  /* ROWS */
  /* ------------------------------ */

  for(let i=0;i<projectData.length;i++){

    const rowNumber = i + 1;

    const item = projectData[i];

    /* ------------------------------ */
    /* ROW HEIGHT */
    /* ------------------------------ */

    /*
      20cm ≈ 567 points
    */

    worksheet.getRow(rowNumber).height = 567;

    /* ------------------------------ */
    /* NOTE */
    /* ------------------------------ */

    const noteCell = worksheet.getCell(`B${rowNumber}`);

    noteCell.value = item.note || "";

    noteCell.alignment = {
      wrapText:true,
      vertical:"top",
      horizontal:"left"
    };

    /* ------------------------------ */
    /* IMAGE */
    /* ------------------------------ */

    if(item.image){

      const base64Data = item.image;

      const extension = base64Data.substring(
        "data:image/".length,
        base64Data.indexOf(";base64")
      );

      const imageId = workbook.addImage({
        base64:base64Data,
        extension:extension
      });

      /* ------------------------------ */
      /* IMAGE SIZE */
      /* ------------------------------ */

      /*
        6cm width ≈ 227px
      */

      const img = new Image();

      img.src = base64Data;

      await new Promise(resolve => {

        img.onload = ()=>{

          const targetWidth = 227;

          const ratio = targetWidth / img.width;

          const targetHeight = img.height * ratio;

          worksheet.addImage(imageId,{

            tl:{
              col:0,
              row:i
            },

            ext:{
              width:targetWidth,
              height:targetHeight
            },

            editAs:"oneCell"
          });

          resolve();
        };

      });

    }

  }

  /* ------------------------------ */
  /* GENERATE FILE */
  /* ------------------------------ */

  const buffer = await workbook.xlsx.writeBuffer();

  saveAs(
    new Blob([buffer]),
    "Image_Notes.xlsx"
  );

});

/* ------------------------------ */
/* STARTUP */
/* ------------------------------ */

window.onload = ()=>{

  const saved = localStorage.getItem("exceljs_project");

  if(!saved){

    continueBtn.style.display = "none";
  }

};