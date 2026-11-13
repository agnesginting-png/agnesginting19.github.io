<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Kuis Informatika Kelas X: Excel dan PowerPoint</title>
<style>
:root{--bg:#f2f6f9;--grid:#d9e3ec;--card:#fff;--ink:#14213d;--muted:#55637a;--line:#c3d0dc;--ok:#0b7a5f;--bad:#b3261e;--accent:#e0a100;--sel:#fff4d1;--btn:#14213d;--btntxt:#fff}
@media (prefers-color-scheme:dark){:root{--bg:#0f1826;--grid:#1a2840;--card:#16233a;--ink:#e8eef7;--muted:#9fb0c8;--line:#2e4261;--ok:#5fd3b0;--bad:#ff8a80;--accent:#f2c14e;--sel:#3a3214;--btn:#e8eef7;--btntxt:#0f1826}}
*{box-sizing:border-box}
html,body{margin:0}
body{background:var(--bg);background-image:linear-gradient(var(--grid) 1px,transparent 1px),linear-gradient(90deg,var(--grid) 1px,transparent 1px);background-size:28px 28px;color:var(--ink);font:16px/1.55 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;min-height:100vh;padding:28px 16px 48px}
main{max-width:680px;margin:0 auto}
header{display:flex;align-items:center;gap:18px;margin-bottom:22px}
h1{font:700 1.8rem/1.15 Georgia,"Iowan Old Style","Times New Roman",serif;margin:0}
header p{margin:4px 0 0;color:var(--muted);font-size:.95rem}
.ring{position:relative;width:76px;height:76px;flex:none}
.ring svg{width:100%;height:100%;transform:rotate(-90deg)}
.ring .t{fill:none;stroke:var(--line);stroke-width:6}
.ring .a{fill:none;stroke:var(--accent);stroke-width:6;stroke-linecap:round;stroke-dasharray:175.93;stroke-dashoffset:175.93;transition:stroke-dashoffset .4s ease}
.ring span{position:absolute;inset:0;display:grid;place-items:center;font-weight:700;font-size:.9rem}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:24px}
.no{color:var(--muted);font-size:.9rem;margin:0 0 6px}
h2{font:600 1.25rem/1.4 Georgia,"Iowan Old Style","Times New Roman",serif;margin:0 0 18px}
label{display:block;font-weight:600;margin:0 0 14px}
input[type=text]{display:block;width:100%;margin-top:6px;font:inherit;color:inherit;background:var(--bg);border:1.5px solid var(--line);border-radius:10px;padding:11px 14px}
ul.info{margin:0 0 18px;padding-left:20px;color:var(--muted);font-size:.95rem}
.opts{display:grid;gap:10px}
.opt{display:flex;gap:12px;align-items:center;text-align:left;width:100%;font:inherit;color:inherit;background:transparent;border:1.5px solid var(--line);border-radius:10px;padding:12px 14px;cursor:pointer}
.opt:hover{border-color:var(--ink)}
.opt .l{flex:none;width:28px;height:28px;border-radius:50%;border:1.5px solid currentColor;display:grid;place-items:center;font-weight:700;font-size:.85rem}
.opt[aria-pressed=true]{border-color:var(--accent);background:var(--sel);border-width:2px}
:focus-visible{outline:3px solid var(--accent);outline-offset:2px}
.nav{display:flex;justify-content:space-between;gap:10px;margin-top:20px}
.btn{font:600 1rem system-ui,sans-serif;background:var(--btn);color:var(--btntxt);border:0;border-radius:10px;padding:12px 22px;cursor:pointer}
.btn.alt{background:transparent;color:var(--ink);border:1.5px solid var(--line)}
.btn:disabled{opacity:.4;cursor:default}
.score{font:700 3.4rem/1 Georgia,serif;margin:4px 0 6px}
.who{margin:0 0 14px;color:var(--muted)}
.who b{color:var(--ink)}
.stat{display:flex;gap:22px;flex-wrap:wrap;margin:10px 0 6px}
.stat div{border:1px solid var(--line);border-radius:10px;padding:8px 16px}
.stat b{display:block;font-size:1.3rem}
.rev{list-style:none;padding:0;margin:18px 0 0;display:grid;gap:10px}
.rev li{border:1px solid var(--line);border-radius:10px;padding:10px 14px;font-size:.95rem}
.rev b{color:var(--bad)}
.rev small{display:block;color:var(--muted)}
.btns{display:flex;gap:10px;flex-wrap:wrap;margin-top:18px}
[hidden]{display:none!important}
@media (prefers-reduced-motion:reduce){.ring .a{transition:none}}
@media print{body{background:#fff;color:#000}.btns,.ring{display:none}.card{border:0;padding:0}}
</style>
</head>
<body>
<main>
  <header>
    <div class="ring" aria-hidden="true">
      <svg viewBox="0 0 76 76"><circle class="t" cx="38" cy="38" r="28"/><circle class="a" id="arc" cx="38" cy="38" r="28"/></svg>
      <span id="cnt">0/30</span>
    </div>
    <div>
      <h1>Kuis Informatika Kelas X</h1>
      <p>Microsoft Excel dan Microsoft PowerPoint, 30 soal pilihan ganda.</p>
    </div>
  </header>

  <section class="card" id="start">
    <h2>Isi identitas dulu</h2>
    <form id="idform">
      <label>Nama lengkap<input type="text" id="nama" required autocomplete="name" maxlength="60"></label>
      <label>Kelas<input type="text" id="kelas" required placeholder="Contoh: X-1" maxlength="20"></label>
      <ul class="info">
        <li>Setiap jawaban benar bernilai 10, nilai maksimal 300.</li>
        <li>Pilih satu jawaban per soal. Kamu bisa kembali dan mengubah jawaban sebelum menekan Selesai.</li>
      </ul>
      <button class="btn" type="submit">Mulai kuis</button>
    </form>
  </section>

  <section class="card" id="quiz" hidden>
    <p class="no" id="qno"></p>
    <h2 id="qtext"></h2>
    <div class="opts" id="opts"></div>
    <div class="nav">
      <button class="btn alt" id="prev">Sebelumnya</button>
      <button class="btn" id="next">Berikutnya</button>
    </div>
  </section>

  <section class="card" id="result" hidden>
    <p class="who">Nama: <b id="rn"></b><br>Kelas: <b id="rk"></b><br>Tanggal: <b id="rt"></b></p>
    <p class="no">Skor akhir</p>
    <div class="score" id="score"></div>
    <p id="msg"></p>
    <div class="stat">
      <div><b id="sb"></b>Benar</div>
      <div><b id="ss"></b>Salah</div>
      <div><b id="sk"></b>Tidak dijawab</div>
    </div>
    <h2 style="margin-top:22px" id="rvh">Soal yang perlu dipelajari lagi</h2>
    <ul class="rev" id="rev"></ul>
    <div class="btns">
      <button class="btn" id="print">Cetak atau simpan PDF</button>
      <button class="btn alt" id="again">Ulangi kuis</button>
    </div>
  </section>
</main>

<script>
const Q=[
// Microsoft Excel
{q:"Ekstensi file default untuk buku kerja (workbook) Microsoft Excel versi modern adalah ....",o:[".docx",".xlsx",".pptx",".txt"],a:1},
{q:"Perpotongan antara baris dan kolom pada lembar kerja Excel disebut ....",o:["Sel (cell)","Range","Sheet","Tabel"],a:0},
{q:"Rumus yang tepat untuk menjumlahkan isi sel A1 sampai A10 adalah ....",o:["=TOTAL(A1:A10)","=SUM(A1:A10)","=ADD(A1:A10)","=PLUS(A1:A10)"],a:1},
{q:"Fungsi Excel untuk menghitung nilai rata-rata sekumpulan data adalah ....",o:["=MEAN()","=AVG()","=AVERAGE()","=MEDIAN()"],a:2},
{q:"Fungsi Excel untuk mencari nilai tertinggi dari sekumpulan data adalah ....",o:["=MIN()","=TOP()","=HIGH()","=MAX()"],a:3},
{q:"Setiap rumus atau fungsi di Excel diawali dengan tanda ....",o:["= (sama dengan)","# (pagar)","$ (dolar)","@ (at)"],a:0},
{q:"Alamat sel yang terletak pada kolom C baris ke-5 ditulis ....",o:["5C","C-5","C5","CR5"],a:2},
{q:"Fungsi Excel untuk menghitung banyaknya sel yang berisi angka adalah ....",o:["=SUM()","=COUNT()","=AVERAGE()","=MAX()"],a:1},
{q:"Fungsi IF pada Excel digunakan untuk ....",o:["menjumlahkan data","mengurutkan data","membuat grafik","menentukan hasil berdasarkan suatu kondisi"],a:3},
{q:"Penulisan alamat sel dengan referensi absolut yang benar adalah ....",o:["A1","$A$1","A1$","[A1]"],a:1},
{q:"Fitur yang digunakan untuk menggabungkan beberapa sel menjadi satu dan meratakan teks di tengah adalah ....",o:["Wrap Text","Format Painter","Merge & Center","Freeze Panes"],a:2},
{q:"Untuk menampilkan data dalam bentuk visual seperti diagram batang, garis, atau lingkaran, kita menggunakan fitur ....",o:["Chart","Filter","Sort","Validation"],a:0},
{q:"Fitur Excel untuk mengurutkan data dari terkecil ke terbesar atau sebaliknya adalah ....",o:["Filter","Sort","Merge","Find"],a:1},
{q:"Satu lembar kerja di dalam file Excel disebut ....",o:["Workbook","Cell","Worksheet","Ribbon"],a:2},
{q:"Fitur Excel untuk menampilkan hanya data yang memenuhi kriteria tertentu adalah ....",o:["Filter","Sort","Merge","Wrap Text"],a:0},
// Microsoft PowerPoint
{q:"Ekstensi file default untuk presentasi Microsoft PowerPoint versi modern adalah ....",o:[".xlsx",".docx",".pptx",".pdf"],a:2},
{q:"Satu halaman pada sebuah presentasi PowerPoint disebut ....",o:["Slide","Sheet","Cell","Frame"],a:0},
{q:"Tombol keyboard untuk memulai tayangan slide dari awal adalah ....",o:["F1","F5","F7","F12"],a:1},
{q:"Efek perpindahan dari satu slide ke slide berikutnya diatur pada tab ....",o:["Animations","Design","Insert","Transitions"],a:3},
{q:"Efek gerak yang diberikan pada objek seperti teks atau gambar di dalam slide diatur pada tab ....",o:["Animations","Transitions","Design","View"],a:0},
{q:"Untuk menyisipkan gambar, tabel, bentuk, atau video ke dalam slide, kita menggunakan tab ....",o:["Home","Design","Insert","View"],a:2},
{q:"Tab yang digunakan untuk mengubah tema dan warna latar seluruh slide adalah ....",o:["Review","Design","Slide Show","Transitions"],a:1},
{q:"Slide Master digunakan untuk ....",o:["menambah animasi pada satu objek","mengecek ejaan teks","mengatur tata letak dan format seluruh slide sekaligus","mencetak catatan pembicara"],a:2},
{q:"Tampilan yang digunakan untuk melihat dan mengatur urutan slide dalam bentuk kumpulan thumbnail adalah ....",o:["Slide Sorter","Notes Page","Outline View","Reading View"],a:0},
{q:"Prinsip penyajian teks yang baik pada slide presentasi adalah ....",o:["teks sebanyak mungkin agar lengkap","singkat, padat, dan mudah dibaca","memakai banyak jenis huruf berbeda","warna mencolok tanpa memperhatikan kontras"],a:1},
{q:"Catatan untuk pembicara yang tidak tampil di layar saat presentasi ditulis pada bagian ....",o:["Slide Master","Ribbon","Status bar","Notes pane"],a:3},
{q:"Kombinasi tombol untuk menyimpan presentasi adalah ....",o:["Ctrl + P","Ctrl + O","Ctrl + S","Ctrl + N"],a:2},
{q:"Layout Title Slide biasanya digunakan untuk ....",o:["menampilkan ringkasan di akhir presentasi","slide judul pada awal presentasi","menampilkan dua kolom perbandingan","menampilkan gambar saja"],a:1},
{q:"Fitur yang digunakan untuk menghubungkan teks atau objek ke slide lain maupun situs web adalah ....",o:["Hyperlink","Animation Pane","Table","Shapes"],a:0},
{q:"Format file presentasi yang langsung berjalan sebagai tayangan slide saat dibuka adalah ....",o:[".docx",".pptx",".xlsx",".ppsx"],a:3}
];
const N=Q.length,PTS=10,$=s=>document.querySelector(s),C=2*Math.PI*28;
let i=0,ans=[],who={};

function ring(n){$('#arc').style.strokeDashoffset=C*(1-n/N);$('#cnt').textContent=n+'/'+N}
function answered(){return ans.filter(v=>v!==undefined).length}

function show(){
  const q=Q[i];
  $('#qno').textContent='Soal '+(i+1)+' dari '+N+(i<15?' (Excel)':' (PowerPoint)');
  $('#qtext').textContent=q.q;
  const box=$('#opts');box.innerHTML='';
  q.o.forEach((t,k)=>{
    const b=document.createElement('button');b.className='opt';b.type='button';
    b.setAttribute('aria-pressed',ans[i]===k);
    const l=document.createElement('span');l.className='l';l.textContent='ABCD'[k];
    const s=document.createElement('span');s.textContent=t;
    b.append(l,s);
    b.onclick=()=>{ans[i]=k;show()};
    box.appendChild(b);
  });
  $('#prev').disabled=i===0;
  $('#next').textContent=i===N-1?'Selesai':'Berikutnya';
  ring(answered());
}

function finish(){
  const kosong=N-answered();
  if(kosong>0&&!confirm('Masih ada '+kosong+' soal yang belum dijawab. Tetap selesaikan kuis?'))return;
  let benar=0;
  Q.forEach((q,k)=>{if(ans[k]===q.a)benar++});
  const salah=answered()-benar;
  $('#quiz').hidden=true;$('#result').hidden=false;
  $('#rn').textContent=who.nama;$('#rk').textContent=who.kelas;
  $('#rt').textContent=new Date().toLocaleDateString('id-ID',{day:'numeric',month:'long',year:'numeric'});
  $('#score').textContent=(benar*PTS)+' / '+(N*PTS);
  $('#sb').textContent=benar;$('#ss').textContent=salah;$('#sk').textContent=kosong;
  const p=benar/N*100;
  $('#msg').textContent=p===100?'Sempurna! Semua jawabanmu benar.':p>=80?'Hebat, pemahamanmu sudah sangat baik.':p>=60?'Cukup baik. Pelajari lagi soal yang salah di bawah.':'Ayo belajar lagi. Pelajari materinya, lalu coba ulang kuis ini.';
  const ul=$('#rev');ul.innerHTML='';
  let n=0;
  Q.forEach((q,k)=>{
    if(ans[k]===q.a)return;
    n++;
    const li=document.createElement('li');
    const b=document.createElement('b');b.textContent='Soal '+(k+1)+'. ';
    const t=document.createTextNode(q.q);
    const s=document.createElement('small');
    s.textContent=(ans[k]===undefined?'Tidak dijawab.':'Jawabanmu: '+q.o[ans[k]]+'.')+' Jawaban benar: '+q.o[q.a]+'.';
    li.append(b,t,s);ul.appendChild(li);
  });
  $('#rvh').hidden=n===0;
  ring(answered());
  window.scrollTo({top:0,behavior:'smooth'});
}

$('#idform').onsubmit=e=>{
  e.preventDefault();
  who={nama:$('#nama').value.trim(),kelas:$('#kelas').value.trim()};
  if(!who.nama||!who.kelas)return;
  $('#start').hidden=true;$('#quiz').hidden=false;
  i=0;ans=[];show();
};
$('#prev').onclick=()=>{if(i>0){i--;show()}};
$('#next').onclick=()=>{if(i<N-1){i++;show()}else finish()};
$('#print').onclick=()=>window.print();
$('#again').onclick=()=>{
  $('#result').hidden=true;$('#start').hidden=false;
  $('#idform').reset();ans=[];ring(0);
  window.scrollTo({top:0,behavior:'smooth'});
};
</script>
</body>
</html>
