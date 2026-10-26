(function(){
  var $ = function(id){return document.getElementById(id)};
  var S = {A:null, B:null};
  var last = null, MAX = 5000;

  function colName(i){var s="";i++;while(i>0){var m=(i-1)%26;s=String.fromCharCode(65+m)+s;i=Math.floor((i-1)/26)}return s}
  function msg(t){$("msg").textContent = t || ""}

  function load(side, file){
    if(!file) return;
    msg("");
    var r = new FileReader();
    r.onload = function(e){
      try{
        var wb = XLSX.read(e.target.result, {type:"array"});
        S[side] = wb;
        $("name"+side).textContent = file.name + " (" + wb.SheetNames.length + " sheet" + (wb.SheetNames.length>1?"s":"") + ")";
        $("drop"+side).classList.add("loaded");
        var sel = $("sheet"+side);
        sel.innerHTML = "";
        wb.SheetNames.forEach(function(n){var o=document.createElement("option");o.value=n;o.textContent=n;sel.appendChild(o)});
        sel.hidden = false;
        refreshKeys();
        $("run").disabled = !(S.A && S.B);
      }catch(err){
        msg("Unable to read this file. Make sure it is a valid Excel or CSV file.");
      }
    };
    r.readAsArrayBuffer(file);
  }

  function rows(side){
    var wb = S[side], ws = wb.Sheets[$("sheet"+side).value];
    return XLSX.utils.sheet_to_json(ws, {header:1, defval:"", raw:false, blankrows:true});
  }

  function refreshKeys(){
    if(!S.A) return;
    var head = rows("A")[0] || [], k = $("key"), cur = k.value;
    k.innerHTML = "";
    var n = Math.max(head.length, 1);
    for(var i=0;i<n;i++){
      var o = document.createElement("option");
      o.value = i;
      o.textContent = colName(i) + (head[i] !== undefined && head[i] !== "" ? " – " + head[i] : "");
      k.appendChild(o);
    }
    if(cur) k.value = cur;
  }

  ["A","B"].forEach(function(s){
    $("file"+s).addEventListener("change", function(e){load(s, e.target.files[0])});
    var d = $("drop"+s);
    d.addEventListener("dragover", function(e){e.preventDefault();d.classList.add("over")});
    d.addEventListener("dragleave", function(){d.classList.remove("over")});
    d.addEventListener("drop", function(e){e.preventDefault();d.classList.remove("over");load(s, e.dataTransfer.files[0])});
    $("sheet"+s).addEventListener("change", function(){if(s==="A") refreshKeys()});
  });
  $("mode").addEventListener("change", function(){$("keyWrap").hidden = this.value !== "key"});

  function norm(v){
    var s = v === undefined || v === null ? "" : String(v);
    if($("trim").checked) s = s.trim();
    if($("icase").checked) s = s.toLowerCase();
    return s;
  }

  function cmpRow(a, b, n){
    var d = [];
    for(var i=0;i<n;i++) if(norm(a[i]) !== norm(b[i])) d.push(i);
    return d;
  }

  function compare(){
    var A = rows("A"), B = rows("B"), out = [];
    var n = 0;
    A.concat(B).forEach(function(r){if(r.length>n) n=r.length});
    if($("mode").value === "pos"){
      var len = Math.max(A.length, B.length);
      for(var i=0;i<len;i++){
        var a = A[i], b = B[i];
        if(a && !b) out.push({st:"removed", a:a, ia:i+1, d:[]});
        else if(!a && b) out.push({st:"added", b:b, ib:i+1, d:[]});
        else{var d = cmpRow(a,b,n); out.push({st:d.length?"modified":"same", a:a, b:b, ia:i+1, ib:i+1, d:d})}
      }
    }else{
      var k = parseInt($("key").value, 10) || 0, map = {}, used = {};
      for(var j=1;j<B.length;j++){
        var kk = norm(B[j][k]), c = 0, key = kk;
        while(map[key] !== undefined){c++; key = kk + "\u0000" + c}
        map[key] = j;
      }
      if(A[0] || B[0]){
        var h1 = A[0] || [], h2 = B[0] || [], dh = cmpRow(h1,h2,n);
        out.push({st:dh.length?"modified":"same", a:h1, b:h2, ia:1, ib:1, d:dh});
      }
      var seen = {};
      for(var x=1;x<A.length;x++){
        var base = norm(A[x][k]), cnt = 0, kx = base;
        while(seen[kx]){cnt++; kx = base + "\u0000" + cnt}
        seen[kx] = true;
        var jb = map[kx];
        if(jb === undefined) out.push({st:"removed", a:A[x], ia:x+1, d:[]});
        else{used[jb] = true; var dd = cmpRow(A[x],B[jb],n); out.push({st:dd.length?"modified":"same", a:A[x], b:B[jb], ia:x+1, ib:jb+1, d:dd})}
      }
      for(var y=1;y<B.length;y++) if(!used[y]) out.push({st:"added", b:B[y], ib:y+1, d:[]});
    }
    last = {rows: out, n: n, head: A[0] || []};
    render();
  }

  function cell(tr, text, cls){
    var td = document.createElement("td");
    if(cls) td.className = cls;
    td.textContent = text;
    tr.appendChild(td);
    return td;
  }


  function plural(n, s, p){return n + " " + (n > 1 ? p : s)}

  function headRow(tbl, labels){
    var tr = document.createElement("tr");
    labels.forEach(function(t){var th = document.createElement("th"); th.textContent = t; tr.appendChild(th)});
    var thead = document.createElement("thead"); thead.appendChild(tr); tbl.appendChild(thead);
  }

  function colLabel(i){
    var h = last.head[i];
    return "Column " + colName(i) + (h !== undefined && h !== "" ? " (" + h + ")" : "");
  }

  function show0(v){return v === "" || v === undefined || v === null ? "(empty)" : v}

  function renderDetails(rowsAll, cells){
    var list = [], byCol = {}, nRows = 0;
    rowsAll.forEach(function(r){
      if(r.st === "modified" && r.d.length){
        nRows++;
        r.d.forEach(function(i){list.push({r:r, i:i}); byCol[i] = (byCol[i] || 0) + 1});
      }
    });
    $("total").textContent = cells
      ? plural(cells, "modified cell", "modified cells") + " in " + plural(nRows, "row", "rows")
      : "No modified cells";
    $("detailWrap").hidden = !cells;
    var ct = $("colTbl"); ct.innerHTML = "";
    headRow(ct, ["Column", "Modified cells"]);
    var cb = document.createElement("tbody");
    Object.keys(byCol).map(Number).sort(function(a,b){return a-b}).forEach(function(i){
      var tr = document.createElement("tr");
      cell(tr, colLabel(i)); cell(tr, byCol[i]);
      cb.appendChild(tr);
    });
    ct.appendChild(cb);
    var dt = $("detTbl"); dt.innerHTML = "";
    headRow(dt, ["Cell", "Row", "Column", "Old value (A)", "New value (B)"]);
    var db = document.createElement("tbody");
    list.slice(0, MAX).forEach(function(x){
      var r = x.r, tr = document.createElement("tr");
      var rowTxt = r.ia === r.ib ? String(r.ia) : "A: " + r.ia + ", B: " + r.ib;
      cell(tr, colName(x.i) + r.ia, "ref");
      cell(tr, rowTxt);
      cell(tr, colLabel(x.i));
      cell(tr, show0(r.a[x.i]), "oldv");
      cell(tr, show0(r.b[x.i]), "newv");
      db.appendChild(tr);
    });
    dt.appendChild(db);
  }

  function render(){
    if(!last) return;
    var cells = 0, added = 0, removed = 0;
    last.rows.forEach(function(r){
      cells += r.d.length;
      if(r.st === "added") added++;
      if(r.st === "removed") removed++;
    });
    renderDetails(last.rows, cells);
    var extra = [];
    if(added) extra.push(plural(added, "row", "rows") + " added in B");
    if(removed) extra.push(plural(removed, "row", "rows") + " missing from B");
    if(extra.length) $("total").textContent += ". " + extra.join(", ") + ".";
    $("result").hidden = false;
  }

  $("run").addEventListener("click", function(){
    msg("");
    try{compare()}catch(err){msg("The comparison failed: " + err.message)}
  });
})();
