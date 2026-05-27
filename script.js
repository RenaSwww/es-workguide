let uploadedData = [];

let customData =
JSON.parse(localStorage.getItem("customData") || "[]");

const stages = {

  beginner:{
    AC:300,
    SM:300,
    TE:300,
    members:3
  },

  normal:{
    AC:500,
    SM:500,
    TE:500,
    members:4
  },

  hard:{
    AC:800,
    SM:800,
    TE:800,
    members:5
  },

  expert:{
    AC:1000,
    SM:900,
    TE:1100,
    members:5
  }

};

document
.getElementById("excelFile")
.addEventListener("change",handleFile);

document
.getElementById("calculateBtn")
.addEventListener("click",calculateTeams);

document
.getElementById("loginBtn")
.addEventListener("click",adminLogin);

document
.getElementById("addDataBtn")
.addEventListener("click",saveCharacter);

document
.getElementById("stageSelect")
.addEventListener("change",changeStage);

function changeStage(e){

  const stage = stages[e.target.value];

  document.getElementById("targetAC").value = stage.AC;
  document.getElementById("targetSM").value = stage.SM;
  document.getElementById("targetTE").value = stage.TE;
  document.getElementById("maxMembers").value = stage.members;
}

function handleFile(e){

  const file = e.target.files[0];

  const reader = new FileReader();

  reader.onload = function(evt){

    const data = new Uint8Array(evt.target.result);

    const workbook =
    XLSX.read(data,{type:"array"});

    const sheet =
    workbook.Sheets[workbook.SheetNames[0]];

    uploadedData =
    XLSX.utils.sheet_to_json(sheet);

    alert("โหลดไฟล์สำเร็จ ✨");
  };

  reader.readAsArrayBuffer(file);
}

function adminLogin(){

  const pass =
  document.getElementById("adminPassword").value;

  if(pass === "admin123"){

    document.getElementById("adminArea")
    .style.display = "block";

    alert("เข้าสู่ระบบสำเร็จ");

  }else{

    alert("รหัสผ่านผิด");
  }
}

function saveCharacter(){

  const newData = {

    Name:
    document.getElementById("charName").value,

    Outfit:
    document.getElementById("outfitName").value,

    AC:
    Number(document.getElementById("acValue").value),

    SM:
    Number(document.getElementById("smValue").value),

    TE:
    Number(document.getElementById("teValue").value),

    Image:
    document.getElementById("imageUrl").value,

    Status:
    document.getElementById("statusValue").value
  };

  customData.push(newData);

  localStorage.setItem(
    "customData",
    JSON.stringify(customData)
  );

  alert("บันทึกข้อมูลสำเร็จ 💾");
}

function calculateTeams(){

  const targetAC =
  Number(document.getElementById("targetAC").value);

  const targetSM =
  Number(document.getElementById("targetSM").value);

  const targetTE =
  Number(document.getElementById("targetTE").value);

  const maxMembers =
  Number(document.getElementById("maxMembers").value);

  const teamCount =
  Number(document.getElementById("teamCount").value);

  const excelData = uploadedData.map(r => ({

    name:r.Name || "Unknown",

    outfit:
    r["Outfit.1"] || r.Outfit || "None",

    status:
    (r.EN || r.Status || "").toLowerCase(),

    AC:
    (Number(r.Active)||0)
    +(Number(r["Active.1"])||0),

    SM:
    (Number(r.Smart)||0)
    +(Number(r["Smart.1"])||0),

    TE:
    (Number(r.Technique)||0)
    +(Number(r["Technique.1"])||0),

    image:
    r.Image ||
    "https://placehold.co/300x300?text=Card"

  }));

  const adminData = customData.map(r => ({

    name:r.Name,
    outfit:r.Outfit,
    status:r.Status,
    AC:r.AC,
    SM:r.SM,
    TE:r.TE,
    image:r.Image

  }));

  const usable =
  [...excelData,...adminData]
  .filter(r =>
    r.status === "done"
    || r.status === "yes"
  );

  let validTeams = [];

  for(let size=1; size<=maxMembers; size++){

    const combos =
    getCombinations(usable,size);

    for(const team of combos){

      const unique =
      new Set(team.map(x=>x.name));

      if(unique.size !== team.length)
      continue;

      const ac =
      team.reduce((s,x)=>s+x.AC,0);

      const sm =
      team.reduce((s,x)=>s+x.SM,0);

      const te =
      team.reduce((s,x)=>s+x.TE,0);

      if(
        ac>=targetAC
        && sm>=targetSM
        && te>=targetTE
      ){

        validTeams.push({

          members:team,
          ac,
          sm,
          te,
          score:ac+sm+te

        });
      }
    }
  }

  validTeams.sort((a,b)=>b.score-a.score);

  renderTeams(
    validTeams.slice(0,teamCount),
    targetAC,
    targetSM,
    targetTE
  );
}

function renderTeams(
  teams,
  targetAC,
  targetSM,
  targetTE
){

  const results =
  document.getElementById("results");

  if(teams.length===0){

    results.innerHTML = `
      <div class="team-card">
        <h2 class="fail">
          ❌ ไม่พบทีม
        </h2>
      </div>
    `;

    return;
  }

  let html =
  '<div class="team-wrapper">';

  teams.forEach((team,index)=>{

    html += `
      <div class="team-card">

      <h2 class="success">
        🏆 Team ${index+1}
      </h2>
    `;

    team.members.forEach(m=>{

      html += `
        <div class="member">

          <img src="${m.image}">

          <div class="member-info">

            <h3>${m.name}</h3>

            <div>
              🎽 ${m.outfit}
            </div>

            <div class="stats">

              <div class="stat">
                AC ${m.AC}
              </div>

              <div class="stat">
                SM ${m.SM}
              </div>

              <div class="stat">
                TE ${m.TE}
              </div>

            </div>

          </div>

        </div>
      `;
    });

    html += `
      <div class="total">

        <div>
          🔥 Active:
          ${team.ac} / ${targetAC}
        </div>

        <div>
          🧠 Smart:
          ${team.sm} / ${targetSM}
        </div>

        <div>
          ⚙️ Technique:
          ${team.te} / ${targetTE}
        </div>

      </div>

      </div>
    `;
  });

  html += '</div>';

  results.innerHTML = html;
}

function getCombinations(arr,size){

  const result = [];

  function helper(start,combo){

    if(combo.length===size){

      result.push([...combo]);

      return;
    }

    for(let i=start;i<arr.length;i++){

      combo.push(arr[i]);

      helper(i+1,combo);

      combo.pop();
    }
  }

  helper(0,[]);

  return result;
}
