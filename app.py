"""
Pesquisa de Preços — IFAL
=========================
Aplicação standalone para realização de pesquisa de preços conforme
Lei nº 14.133/2021 e IN SEGES/ME nº 65/2021.

COMO USAR:
1. Instale as dependências:
   pip install flask anthropic openpyxl reportlab

2. Configure sua chave de API da Anthropic:
   Windows:   set ANTHROPIC_API_KEY=sk-ant-...
   Mac/Linux: export ANTHROPIC_API_KEY=sk-ant-...

3. Execute o servidor:
   python app.py

4. Acesse no navegador:
   http://localhost:5000
"""

import os, json, io, re
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template_string

app = Flask(__name__)

# ─── HTML Frontend ───────────────────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Pesquisa de Preços — IFAL</title>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --g:#1B5E20;--gl:#2E7D32;--gp:#E8F5E9;--gm:#4CAF50;
  --acc:#FF8F00;--accl:#FFF8E1;
  --bg:#F7F8F5;--surf:#fff;--bdr:#DDE3D8;--bdrs:#B0BDA8;
  --tx:#1A2116;--txs:#4A5842;--txm:#7A8A72;
  --err:#C62828;--errl:#FFEBEE;
  --r:8px;--rl:12px;
  --sh:0 1px 3px rgba(0,0,0,.08),0 1px 2px rgba(0,0,0,.04);
}
body{font-family:'IBM Plex Sans',sans-serif;background:var(--bg);color:var(--tx);min-height:100vh;font-size:15px;line-height:1.6}
header{background:var(--g);color:#fff;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.2)}
.hdr{max-width:1100px;margin:0 auto;padding:14px 24px;display:flex;align-items:center;gap:16px}
.logo{width:44px;height:44px;background:#fff;border-radius:50%;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.logo svg{width:32px;height:32px}
.htxt h1{font-size:17px;font-weight:600}
.htxt p{font-size:12px;opacity:.75;margin-top:1px}
.badge{margin-left:auto;background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);border-radius:20px;padding:4px 12px;font-size:11px;font-weight:500;letter-spacing:.5px;text-transform:uppercase}
main{max-width:1100px;margin:0 auto;padding:32px 24px 80px}
.card{background:var(--surf);border:1px solid var(--bdr);border-radius:var(--rl);padding:24px;margin-bottom:20px;box-shadow:var(--sh)}
.stitle{font-size:13px;font-weight:600;text-transform:uppercase;letter-spacing:.8px;color:var(--g);margin-bottom:18px;display:flex;align-items:center;gap:8px}
.stitle::after{content:'';flex:1;height:1px;background:var(--gp)}
.frow{display:grid;grid-template-columns:1fr auto auto auto;gap:10px;align-items:end}
.fg{display:flex;flex-direction:column;gap:5px}
.fg label{font-size:12px;font-weight:500;color:var(--txs)}
input,select{height:38px;padding:0 12px;border:1px solid var(--bdr);border-radius:var(--r);font-family:inherit;font-size:14px;color:var(--tx);background:var(--surf);outline:none;transition:border-color .15s,box-shadow .15s}
input:focus,select:focus{border-color:var(--gm);box-shadow:0 0 0 3px rgba(76,175,80,.12)}
.btn{height:38px;padding:0 18px;border-radius:var(--r);border:none;font-family:inherit;font-size:13px;font-weight:500;cursor:pointer;display:inline-flex;align-items:center;gap:7px;transition:all .15s;white-space:nowrap}
.bp{background:var(--g);color:#fff}.bp:hover{background:var(--gl)}.bp:disabled{opacity:.55;cursor:not-allowed}
.bs{background:transparent;color:var(--g);border:1px solid var(--g)}.bs:hover{background:var(--gp)}
.ba{background:var(--acc);color:#fff}.ba:hover{background:#F57F00}
.btn-sm{height:30px;padding:0 12px;font-size:12px}
.dz{border:2px dashed var(--bdrs);border-radius:var(--rl);padding:20px;text-align:center;cursor:pointer;transition:all .2s;color:var(--txm);font-size:13px}
.dz:hover,.dz.over{border-color:var(--gm);background:var(--gp);color:var(--g)}
.dz p{margin-top:6px;font-size:12px}
.twrap{overflow-x:auto;margin-top:4px}
table{width:100%;border-collapse:collapse;font-size:13px}
thead th{text-align:left;padding:8px 10px;background:var(--bg);border-bottom:1px solid var(--bdr);font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.6px;color:var(--txs);white-space:nowrap;position:sticky;top:0}
tbody td{padding:7px 10px;border-bottom:1px solid var(--bdr);vertical-align:middle}
tbody tr:last-child td{border-bottom:none}
tbody tr:hover td{background:var(--gp)}
.ie{border:1px solid transparent;background:transparent;padding:2px 6px;border-radius:4px;font-family:inherit;font-size:13px;color:var(--tx);width:100%;height:28px}
.ie:hover{border-color:var(--bdr)}
.ie:focus{border-color:var(--gm);background:#fff;box-shadow:0 0 0 2px rgba(76,175,80,.1)}
.pi{width:90px;height:28px;font-size:12px;padding:0 8px;font-family:'IBM Plex Mono',monospace}
.pi:focus{border-color:var(--gm);box-shadow:0 0 0 2px rgba(76,175,80,.1)}
.avg{font-family:'IBM Plex Mono',monospace;font-size:13px;font-weight:600;color:var(--g)}
.tot{font-family:'IBM Plex Mono',monospace;font-size:13px;font-weight:600;color:var(--acc)}
.sbadge{display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:500}
.sp{background:var(--accl);color:#8B6000}
.sl{background:#E3F2FD;color:#1565C0}
.sd{background:var(--gp);color:var(--g)}
.se{background:var(--errl);color:var(--err)}
@keyframes spin{to{transform:rotate(360deg)}}
.spin{width:10px;height:10px;border:2px solid currentColor;border-top-color:transparent;border-radius:50%;animation:spin .7s linear infinite;display:inline-block}
.abar{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:16px;padding-top:16px;border-top:1px solid var(--bdr)}
.prog{display:none;margin-top:16px;padding:14px 18px;background:var(--gp);border-radius:var(--r);border:1px solid rgba(76,175,80,.2)}
.prog.on{display:block}
.plab{font-size:13px;color:var(--g);margin-bottom:8px;font-weight:500}
.pbg{height:6px;background:rgba(76,175,80,.2);border-radius:3px;overflow:hidden}
.pfill{height:100%;background:var(--g);border-radius:3px;transition:width .4s ease}
.toast{position:fixed;bottom:24px;right:24px;background:var(--tx);color:#fff;padding:12px 18px;border-radius:var(--r);font-size:13px;box-shadow:0 4px 12px rgba(0,0,0,.15);transform:translateY(20px);opacity:0;transition:all .3s;z-index:999;max-width:360px;pointer-events:none}
.toast.on{transform:translateY(0);opacity:1}
.toast.ok{background:var(--g)}.toast.err{background:var(--err)}
.empty{text-align:center;padding:40px 20px;color:var(--txm)}
.empty .ico{font-size:36px;margin-bottom:10px}
.fnote{font-size:11px;color:var(--txm);margin-top:12px;padding-top:12px;border-top:1px solid var(--bdr);line-height:1.5}
.xbtn{background:transparent;border:none;color:var(--err);font-size:18px;cursor:pointer;padding:0 6px;border-radius:4px;line-height:1}
.xbtn:hover{background:var(--errl)}
@media(max-width:768px){.frow{grid-template-columns:1fr};main{padding:20px 16px 60px}}
</style>
</head>
<body>

<header>
  <div class="hdr">
    <div class="logo">
      <svg viewBox="0 0 44 44"><circle cx="22" cy="22" r="22" fill="#1B5E20"/>
        <text x="22" y="27" font-family="IBM Plex Sans,sans-serif" font-size="11" font-weight="700" fill="white" text-anchor="middle">IFAL</text>
      </svg>
    </div>
    <div class="htxt">
      <h1>Pesquisa de Preços</h1>
      <p>Instituto Federal de Alagoas — Contratações Públicas (Lei nº 14.133/2021)</p>
    </div>
    <div class="badge">IN 65/2021</div>
  </div>
</header>

<main>

  <!-- Adicionar item -->
  <div class="card">
    <div class="stitle">Adicionar Item</div>
    <div class="frow">
      <div class="fg">
        <label>Descrição completa do item *</label>
        <input id="inp-desc" style="width:100%" placeholder="Ex: Papel A4 75g/m², branco, resma com 500 folhas" autocomplete="off">
      </div>
      <div class="fg">
        <label>Unidade</label>
        <select id="inp-unit" style="width:160px">
          <option value="un">Unidade</option>
          <option value="resma">Resma</option>
          <option value="caixa">Caixa</option>
          <option value="pacote">Pacote</option>
          <option value="kg">Kg</option>
          <option value="g">Grama (g)</option>
          <option value="litro">Litro</option>
          <option value="ml">ml</option>
          <option value="metro">Metro</option>
          <option value="par">Par</option>
          <option value="conjunto">Conjunto</option>
          <option value="rolo">Rolo</option>
          <option value="frasco">Frasco</option>
          <option value="galão">Galão</option>
          <option value="diária">Diária</option>
          <option value="m²">m²</option>
          <option value="metro linear">Metro Linear</option>
          <option value="por pessoa">Por Pessoa</option>
          <option value="outro">Outro...</option>
        </select>
      </div>
      <div class="fg">
        <label>Quantidade</label>
        <input id="inp-qty" type="number" min="1" style="width:100px" placeholder="Ex: 100">
      </div>
      <div class="fg">
        <label>&nbsp;</label>
        <button class="btn bp" onclick="addItem()">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
          Adicionar
        </button>
      </div>
    </div>
    <div id="unit-custom" style="display:none;margin-top:10px">
      <input id="inp-unit-custom" style="width:220px" placeholder="Digite a unidade personalizada...">
    </div>
  </div>

  <!-- Importar Excel -->
  <div class="card">
    <div class="stitle">Importar Planilha Excel</div>
    <div class="dz" id="dz" onclick="document.getElementById('fi').click()">
      <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#4CAF50" stroke-width="1.5" style="margin:0 auto;display:block">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>
        <line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/>
      </svg>
      <strong style="display:block;margin-top:8px;font-size:14px;color:#1A2116">Clique ou arraste o arquivo .xlsx aqui</strong>
      <p>Colunas esperadas: <strong>Descrição</strong>, <strong>Unidade</strong>, <strong>Quantidade</strong> (nessa ordem ou com esses nomes)</p>
    </div>
    <input type="file" id="fi" accept=".xlsx,.xls,.csv" style="display:none" onchange="handleFile(this)">
  </div>

  <!-- Tabela de itens -->
  <div class="card">
    <div class="stitle">Itens da Pesquisa</div>

    <div class="empty" id="empty">
      <div class="ico">📋</div>
      <p>Nenhum item adicionado ainda.</p>
      <p style="margin-top:4px">Adicione itens manualmente ou importe uma planilha Excel.</p>
    </div>

    <div class="twrap" id="twrap" style="display:none">
      <table>
        <thead><tr>
          <th>#</th><th style="min-width:200px">Descrição</th><th>Unidade</th><th>Qtd.</th>
          <th>Preço 1 (R$)</th><th>Fonte 1</th>
          <th>Preço 2 (R$)</th><th>Fonte 2</th>
          <th>Preço 3 (R$)</th><th>Fonte 3</th>
          <th>Média (R$)</th><th>Total (R$)</th><th>Status</th><th></th>
        </tr></thead>
        <tbody id="tbody"></tbody>
      </table>
    </div>

    <div class="abar" id="abar" style="display:none">
      <button class="btn bs btn-sm" onclick="clearAll()">
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/>
        </svg>
        Limpar tudo
      </button>
      <div style="flex:1"></div>
      <button class="btn bp" id="btn-research" onclick="startResearch()">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
        </svg>
        Realizar Pesquisa de Preços
      </button>
      <button class="btn ba" id="btn-pdf" style="display:none" onclick="generatePDF()">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
          <polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/>
        </svg>
        Gerar PDF
      </button>
    </div>

    <div class="prog" id="prog">
      <div class="plab" id="plab">Pesquisando…</div>
      <div class="pbg"><div class="pfill" id="pfill" style="width:0%"></div></div>
    </div>

    <p class="fnote">
      ⚠️ A pesquisa de preços é realizada via IA com base em dados de mercado disponíveis.
      Os valores coletados devem ser verificados e validados pelo servidor responsável antes da elaboração do processo licitatório.
      Preços de pessoas físicas e produtos usados em plataformas como Mercado Livre e Amazon não são considerados.
    </p>
  </div>
</main>

<div class="toast" id="toast"></div>

<script>
const items = [];
let researchDone = false;

/* ── toast ── */
function toast(msg, type='') {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className = 'toast on' + (type ? ' '+type : '');
  clearTimeout(el._t);
  el._t = setTimeout(() => el.className='toast', 3500);
}

/* ── currency ── */
function fmtBR(v) {
  if (v==='' || v===null || v===undefined || isNaN(parseFloat(v))) return null;
  return parseFloat(v).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});
}
function calcAvg(it) {
  const vs = [it.price1,it.price2,it.price3].filter(v=>v!==''&&!isNaN(parseFloat(v))).map(Number);
  return vs.length ? vs.reduce((a,b)=>a+b,0)/vs.length : null;
}
function calcTot(it) {
  const avg=calcAvg(it);
  return avg!==null && it.qty ? avg*parseFloat(it.qty) : null;
}

/* ── add item ── */
function addItem() {
  const d = document.getElementById('inp-desc').value.trim();
  const us = document.getElementById('inp-unit').value;
  const u = us==='outro' ? (document.getElementById('inp-unit-custom').value.trim()||'Outro') : us;
  const q = parseFloat(document.getElementById('inp-qty').value);
  if (!d) { toast('Informe a descrição.','err'); return; }
  if (!q||q<=0) { toast('Informe uma quantidade válida.','err'); return; }
  items.push({id:Date.now(), desc:d, unit:u, qty:q, price1:'',src1:'',price2:'',src2:'',price3:'',src3:'',status:'pending'});
  document.getElementById('inp-desc').value='';
  document.getElementById('inp-qty').value='';
  render();
  toast('Item adicionado!','ok');
}

document.getElementById('inp-unit').addEventListener('change',function(){
  document.getElementById('unit-custom').style.display = this.value==='outro'?'block':'none';
});
document.getElementById('inp-desc').addEventListener('keydown',e=>{if(e.key==='Enter')addItem();});

/* ── render ── */
function render() {
  document.getElementById('empty').style.display   = items.length ? 'none' : '';
  document.getElementById('twrap').style.display   = items.length ? '' : 'none';
  document.getElementById('abar').style.display    = items.length ? '' : 'none';
  document.getElementById('btn-pdf').style.display = researchDone ? '' : 'none';

  const tb = document.getElementById('tbody');
  tb.innerHTML='';
  items.forEach((it,idx)=>{
    const avg=calcAvg(it), tot=calcTot(it);
    const avgTxt = avg!==null ? 'R$ '+fmtBR(avg) : '—';
    const totTxt = tot!==null ? 'R$ '+fmtBR(tot) : '—';
    const statusHTML = {
      pending:'<span class="sbadge sp">Aguardando</span>',
      loading:'<span class="sbadge sl"><span class="spin"></span> Pesquisando</span>',
      done:   '<span class="sbadge sd">✓ Concluído</span>',
      error:  '<span class="sbadge se">⚠ Erro</span>',
    }[it.status]||'';

    const tr=document.createElement('tr');
    tr.innerHTML=
      `<td style="color:var(--txm);font-size:12px">${idx+1}</td>`+
      `<td><input class="ie" value="${esc(it.desc)}" data-id="${it.id}" data-f="desc" style="min-width:180px"></td>`+
      `<td><input class="ie" value="${esc(it.unit)}" data-id="${it.id}" data-f="unit" style="width:80px"></td>`+
      `<td><input class="ie" type="number" value="${it.qty}" data-id="${it.id}" data-f="qty" style="width:65px;font-family:'IBM Plex Mono',monospace;font-size:12px"></td>`+
      `<td><input class="ie pi" type="number" value="${it.price1}" placeholder="0,00" step="0.01" data-id="${it.id}" data-f="price1"></td>`+
      `<td><input class="ie" value="${esc(it.src1)}" placeholder="Fonte..." data-id="${it.id}" data-f="src1" style="min-width:110px"></td>`+
      `<td><input class="ie pi" type="number" value="${it.price2}" placeholder="0,00" step="0.01" data-id="${it.id}" data-f="price2"></td>`+
      `<td><input class="ie" value="${esc(it.src2)}" placeholder="Fonte..." data-id="${it.id}" data-f="src2" style="min-width:110px"></td>`+
      `<td><input class="ie pi" type="number" value="${it.price3}" placeholder="0,00" step="0.01" data-id="${it.id}" data-f="price3"></td>`+
      `<td><input class="ie" value="${esc(it.src3)}" placeholder="Fonte..." data-id="${it.id}" data-f="src3" style="min-width:110px"></td>`+
      `<td><span class="avg" id="avg-${it.id}">${avgTxt}</span></td>`+
      `<td><span class="tot" id="tot-${it.id}">${totTxt}</span></td>`+
      `<td>${statusHTML}</td>`+
      `<td><button class="xbtn" data-id="${it.id}" title="Remover">×</button></td>`;
    tb.appendChild(tr);
  });

  tb.querySelectorAll('input[data-f]').forEach(inp=>{
    inp.addEventListener('change',function(){
      const it=items.find(x=>x.id==this.dataset.id);
      if(!it) return;
      const f=this.dataset.f;
      it[f] = f==='qty' ? parseFloat(this.value)||1 : this.value;
      if(['price1','price2','price3','qty'].includes(f)) {
        const a=calcAvg(it), t=calcTot(it);
        const ae=document.getElementById('avg-'+it.id);
        const te=document.getElementById('tot-'+it.id);
        if(ae) ae.textContent = a!==null?'R$ '+fmtBR(a):'—';
        if(te) te.textContent = t!==null?'R$ '+fmtBR(t):'—';
      }
    });
  });
  tb.querySelectorAll('button[data-id]').forEach(b=>{
    b.addEventListener('click',function(){
      const idx=items.findIndex(x=>x.id==this.dataset.id);
      if(idx>=0){items.splice(idx,1); render();}
    });
  });
}

function esc(s){return String(s||'').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function clearAll(){if(!confirm('Remover todos os itens?'))return;items.length=0;researchDone=false;render();}

/* ── Excel ── */
const dz=document.getElementById('dz');
dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('over');});
dz.addEventListener('dragleave',()=>dz.classList.remove('over'));
dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('over');if(e.dataTransfer.files[0])processXL(e.dataTransfer.files[0]);});
function handleFile(inp){if(inp.files[0])processXL(inp.files[0]);inp.value='';}

function processXL(file){
  const r=new FileReader();
  r.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'array'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const data=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
      if(!data.length){toast('Planilha vazia.','err');return;}
      const hdrs=data[0].map(h=>String(h||'').toLowerCase().trim());
      const fc=(vs)=>{
        for(const v of vs){const i=hdrs.indexOf(v);if(i>=0)return i;}
        for(const v of vs){const i=hdrs.findIndex(h=>h.includes(v));if(i>=0)return i;}
        return -1;
      };
      const dI=fc(['descricao','descrição','item','produto','material','desc']);
      const uI=fc(['unidade','un','und','unit','umd','unid']);
      const qI=fc(['quantidade','qtd','qtde','qty','quant']);
      let d=0,u=1,q=2,start=0;
      if(dI>=0){d=dI;u=uI>=0?uI:dI+1;q=qI>=0?qI:u+1;start=1;}
      let cnt=0;
      for(let r2=start;r2<data.length;r2++){
        const row=data[r2];
        const desc=String(row[d]||'').trim();
        const unit=String(row[u]||'un').trim()||'un';
        const qty=parseFloat(String(row[q]||'').replace(',','.'));
        if(!desc||isNaN(qty)||qty<=0)continue;
        items.push({id:Date.now()+r2,desc,unit,qty,price1:'',src1:'',price2:'',src2:'',price3:'',src3:'',status:'pending'});
        cnt++;
      }
      render();
      toast(cnt+' item(ns) importado(s)!','ok');
    }catch(err){console.error(err);toast('Erro ao ler o arquivo.','err');}
  };
  r.readAsArrayBuffer(file);
}

/* ── Pesquisa ── */
async function startResearch(){
  if(!items.length){toast('Adicione itens antes de pesquisar.','err');return;}
  if(!confirm('Iniciar pesquisa de preços para '+items.length+' item(ns)?\nIsso pode levar alguns minutos.'))return;

  const prog=document.getElementById('prog');
  const pfill=document.getElementById('pfill');
  const plab=document.getElementById('plab');
  const btn=document.getElementById('btn-research');
  prog.classList.add('on');
  btn.disabled=true;

  for(let i=0;i<items.length;i++){
    items[i].status='loading';
    render();
    pfill.style.width=Math.round((i/items.length)*100)+'%';
    plab.textContent='Pesquisando item '+(i+1)+' de '+items.length+': '+items[i].desc.substring(0,60)+'…';

    try{
      const res=await fetch('/api/research',{
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body:JSON.stringify({desc:items[i].desc,unit:items[i].unit})
      });
      if(!res.ok){
        const err=await res.json().catch(()=>({}));
        throw new Error(err.error||'HTTP '+res.status);
      }
      const data=await res.json();
      items[i].price1=data.price1; items[i].src1=data.src1;
      items[i].price2=data.price2; items[i].src2=data.src2;
      items[i].price3=data.price3; items[i].src3=data.src3;
      items[i].status='done';
    }catch(err){
      console.error('Item',i,err);
      items[i].status='error';
    }
    render();
    await new Promise(r=>setTimeout(r,200));
  }

  pfill.style.width='100%';
  plab.textContent='✓ Pesquisa concluída!';
  setTimeout(()=>prog.classList.remove('on'),2500);
  btn.disabled=false;
  researchDone=true;
  render();
  toast('Pesquisa concluída!','ok');
}

/* ── Gerar PDF ── */
async function generatePDF(){
  const payload=items.map(it=>({
    desc:it.desc, unit:it.unit, qty:it.qty,
    price1:it.price1, src1:it.src1,
    price2:it.price2, src2:it.src2,
    price3:it.price3, src3:it.src3,
  }));
  toast('Gerando PDF…');
  try{
    const res=await fetch('/api/pdf',{
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({items:payload})
    });
    if(!res.ok){const e=await res.json().catch(()=>({}));throw new Error(e.error||'Erro ao gerar PDF');}
    const blob=await res.blob();
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url;
    const d=new Date().toLocaleDateString('pt-BR').replace(/\//g,'-');
    a.download='pesquisa-precos-ifal-'+d+'.pdf';
    a.click();
    URL.revokeObjectURL(url);
    toast('PDF baixado com sucesso!','ok');
  }catch(err){
    console.error(err);
    toast('Erro ao gerar PDF: '+err.message,'err');
  }
}

render();
</script>
</body>
</html>"""


# ─── API: pesquisa de preços ─────────────────────────────────────────────────

@app.route('/api/research', methods=['POST'])
def research():
    try:
        import anthropic
    except ImportError:
        return jsonify({'error': 'Biblioteca anthropic não instalada. Execute: pip install anthropic'}), 500

    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        return jsonify({'error': 'Variável ANTHROPIC_API_KEY não definida.'}), 500

    data = request.get_json()
    desc = data.get('desc', '')
    unit = data.get('unit', 'un')

    client = anthropic.Anthropic(api_key=api_key)

    system = (
        "Você é um assistente especializado em pesquisa de preços para contratações públicas no Brasil, "
        "seguindo a IN SEGES/ME nº 65/2021 e a Lei nº 14.133/2021.\n\n"
        "Para cada item, forneça EXATAMENTE 3 preços unitários de fontes distintas.\n\n"
        "REGRAS:\n"
        "1. Apenas fornecedores PJ com emissão de NF-e.\n"
        "2. NÃO inclua pessoas físicas, produtos usados ou seminovos.\n"
        "3. Mercado Livre/Amazon: só vendedores PJ com NF-e.\n"
        "4. Priorize: Painel de Preços, ComprasNet/PNCP, depois grandes varejistas.\n"
        "5. Preços realistas para o mercado brasileiro em " + str(datetime.now().year) + ".\n"
        "6. Responda SOMENTE com JSON puro, sem markdown, sem texto adicional.\n\n"
        'Formato: {"price1":12.50,"src1":"Nome Fonte 1","price2":13.80,"src2":"Nome Fonte 2","price3":11.90,"src3":"Nome Fonte 3"}'
    )

    user = (
        f"Pesquise 3 preços unitários para o seguinte item de contratação pública do IFAL:\n"
        f"Descrição: {desc}\n"
        f"Unidade: {unit}\n"
        f"Forneça o valor unitário (por {unit}) em R$ de 3 fornecedores PJ distintos com NF-e."
    )

    message = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=300,
        system=system,
        messages=[{"role": "user", "content": user}]
    )

    raw = "".join(b.text for b in message.content if b.type == "text")
    m = re.search(r'\{[\s\S]*\}', raw)
    if not m:
        return jsonify({'error': 'Resposta inválida da IA: ' + raw[:100]}), 500

    parsed = json.loads(m.group(0))
    return jsonify({
        'price1': float(parsed.get('price1') or 0),
        'src1':   str(parsed.get('src1') or 'Fonte 1')[:80],
        'price2': float(parsed.get('price2') or 0),
        'src2':   str(parsed.get('src2') or 'Fonte 2')[:80],
        'price3': float(parsed.get('price3') or 0),
        'src3':   str(parsed.get('src3') or 'Fonte 3')[:80],
    })


# ─── API: geração de PDF ─────────────────────────────────────────────────────

def fmt_br(v):
    try:
        return f"R$ {float(v):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return '—'

def avg_of(item):
    vals = [item.get(k) for k in ('price1', 'price2', 'price3')]
    vals = [float(v) for v in vals if v not in ('', None) and str(v).strip() != '']
    return sum(vals) / len(vals) if vals else None

def total_of(item):
    avg = avg_of(item)
    qty = item.get('qty')
    try:
        return avg * float(qty) if avg is not None and qty else None
    except:
        return None


@app.route('/api/pdf', methods=['POST'])
def generate_pdf():
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                    Paragraph, Spacer, PageBreak, KeepTogether)
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT

    data = request.get_json()
    items = data.get('items', [])
    date_str = datetime.now().strftime('%d/%m/%Y')
    year_str = datetime.now().strftime('%Y')

    buf = io.BytesIO()
    PAGE_W, PAGE_H = landscape(A4)
    LM = 12 * mm
    RM = 12 * mm
    TM = 30 * mm
    BM = 12 * mm

    # ── colors ──
    GREEN      = colors.HexColor('#1B5E20')
    GREEN_PALE = colors.HexColor('#E8F5E9')
    ORANGE     = colors.HexColor('#FF8F00')
    GREY_BDR   = colors.HexColor('#DDE3D8')
    WHITE      = colors.white
    BLACK      = colors.black
    DARK_GREEN = colors.HexColor('#1B5E20')

    # ── styles ──
    styles = getSampleStyleSheet()
    small = ParagraphStyle('small', fontSize=6.5, leading=8, wordWrap='CJK')
    small_c = ParagraphStyle('smallC', fontSize=6.5, leading=8, alignment=TA_CENTER, wordWrap='CJK')
    small_r = ParagraphStyle('smallR', fontSize=6.5, leading=8, alignment=TA_RIGHT, wordWrap='CJK')
    bold_r  = ParagraphStyle('boldR',  fontSize=6.5, leading=8, alignment=TA_RIGHT,
                              fontName='Helvetica-Bold', textColor=DARK_GREEN, wordWrap='CJK')
    bold_acc= ParagraphStyle('boldAcc',fontSize=6.5, leading=8, alignment=TA_RIGHT,
                              fontName='Helvetica-Bold', textColor=colors.HexColor('#8B6000'), wordWrap='CJK')

    def p(txt, style=None): return Paragraph(str(txt or '—'), style or small)
    def pc(txt): return Paragraph(str(txt or ''), small_c)
    def pr(txt): return Paragraph(str(txt or '—'), small_r)

    # ── grand total ──
    grand_total = sum(total_of(it) or 0 for it in items)

    # ── table rows ──
    col_widths = [7*mm, 52*mm, 12*mm, 10*mm,
                  18*mm, 20*mm, 18*mm, 20*mm, 18*mm, 20*mm,
                  20*mm, 23*mm]

    head_row = [
        pc('#'), p('Descrição do Item', small_c), pc('Unid.'), pc('Qtd.'),
        pc('Valor P1'), pc('Fonte P1'), pc('Valor P2'), pc('Fonte P2'),
        pc('Valor P3'), pc('Fonte P3'), pc('Média'), pc('Valor Total')
    ]

    rows = [head_row]
    for idx, it in enumerate(items):
        avg = avg_of(it)
        tot = total_of(it)
        rows.append([
            pc(str(idx + 1)),
            p(it.get('desc', '')),
            pc(it.get('unit', '')),
            pc(str(it.get('qty', ''))),
            pr(fmt_br(it.get('price1')) if it.get('price1') not in ('', None) else '—'),
            p(it.get('src1') or '—'),
            pr(fmt_br(it.get('price2')) if it.get('price2') not in ('', None) else '—'),
            p(it.get('src2') or '—'),
            pr(fmt_br(it.get('price3')) if it.get('price3') not in ('', None) else '—'),
            p(it.get('src3') or '—'),
            Paragraph(fmt_br(avg) if avg is not None else '—', bold_r),
            Paragraph(fmt_br(tot) if tot is not None else '—', bold_acc),
        ])

    # total geral row
    rows.append([
        pc(''), pc(''), pc(''), pc(''), pc(''), pc(''), pc(''), pc(''), pc(''), pc(''),
        Paragraph('TOTAL GERAL', ParagraphStyle('tg', fontSize=7, fontName='Helvetica-Bold',
                                                 alignment=TA_RIGHT, textColor=WHITE)),
        Paragraph(fmt_br(grand_total), ParagraphStyle('tgv', fontSize=7, fontName='Helvetica-Bold',
                                                        alignment=TA_RIGHT, textColor=WHITE)),
    ])

    total_row_idx = len(rows) - 1

    tbl = Table(rows, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle([
        # Header
        ('BACKGROUND',  (0,0), (-1,0), GREEN),
        ('TEXTCOLOR',   (0,0), (-1,0), WHITE),
        ('FONTNAME',    (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',    (0,0), (-1,0), 6.5),
        ('ALIGN',       (0,0), (-1,0), 'CENTER'),
        ('VALIGN',      (0,0), (-1,-1), 'MIDDLE'),
        ('GRID',        (0,0), (-1,-2), 0.3, GREY_BDR),
        ('ROWBACKGROUNDS', (0,1), (-1,-2), [WHITE, colors.HexColor('#F7FBF7')]),
        ('TOPPADDING',  (0,0), (-1,-1), 3),
        ('BOTTOMPADDING',(0,0),(-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('RIGHTPADDING',(0,0), (-1,-1), 3),
        # Total geral row
        ('BACKGROUND',  (0, total_row_idx), (-1, total_row_idx), ORANGE),
        ('GRID',        (0, total_row_idx), (-1, total_row_idx), 0.3, colors.HexColor('#CC7000')),
        ('SPAN',        (0, total_row_idx), (9, total_row_idx)),
    ]))

    # ── methodology content ──
    msty = getSampleStyleSheet()
    title_sty = ParagraphStyle('mt', fontSize=9, fontName='Helvetica-Bold',
                                textColor=DARK_GREEN, spaceAfter=4)
    body_sty  = ParagraphStyle('mb', fontSize=7.5, leading=11, spaceAfter=3)
    bul_sty   = ParagraphStyle('mbul', fontSize=7.5, leading=11,
                                leftIndent=10, spaceAfter=2,
                                bulletIndent=4, bulletText='\u2022')

    def sec(txt):   return Paragraph(txt, title_sty)
    def para(txt):  return Paragraph(txt, body_sty)
    def bul(txt):   return Paragraph(txt, bul_sty)
    def sp(h=4):    return Spacer(1, h)

    methodology = [
        sec('1. FUNDAMENTAÇÃO LEGAL'),
        para(
            'A presente pesquisa de preços foi realizada em conformidade com a <b>Lei nº 14.133, de 1º de abril de 2021</b> '
            '(Nova Lei de Licitações e Contratos Administrativos), que em seu art. 18, §1º, inciso IV, estabelece a pesquisa '
            'de preços como etapa indispensável da fase preparatória da contratação pública, bem como com a '
            '<b>Instrução Normativa SEGES/ME nº 65, de 7 de julho de 2021</b>, que regulamenta o procedimento '
            'administrativo para a realização de pesquisa de preços no âmbito da Administração Pública Federal.'
        ), sp(),
        sec('2. PARÂMETROS UTILIZADOS (IN 65/2021, art. 5º)'),
        para('A IN SEGES/ME nº 65/2021 estabelece os seguintes parâmetros, em ordem de prioridade:'),
        bul('I — Painel de Preços (paineldeprecos.planejamento.gov.br)'),
        bul('II — Aquisições e contratações similares de outros entes públicos (PNCP, ComprasNet)'),
        bul('III — Dados de pesquisa em mídia especializada, sítios eletrônicos especializados ou de domínio amplo'),
        bul('IV — Pesquisa direta com no mínimo 3 fornecedores'),
        bul('V — Pesquisa em base de notas fiscais eletrônicas'), sp(),
        sec('3. METODOLOGIA DE CÁLCULO'),
        para(
            'Conforme art. 6º da IN SEGES/ME nº 65/2021, o preço estimado foi apurado mediante a <b>média aritmética</b> '
            'dos preços coletados junto a, no mínimo, 3 (três) fontes distintas. Não foram considerados valores '
            'inexequíveis ou manifestamente excessivos.'
        ),
        para('<b>Fórmula:</b> Média = (Preço 1 + Preço 2 + Preço 3) / 3  |  Valor Total = Média × Quantidade'), sp(),
        sec('4. CRITÉRIOS DE EXCLUSÃO DE FONTES'),
        para('Foram adotados os seguintes critérios de exclusão:'),
        bul('Não foram considerados preços coletados de vendedores pessoa física em plataformas de comércio eletrônico (Mercado Livre, Amazon, OLX e similares);'),
        bul('Não foram considerados preços de produtos usados, recondicionados ou seminovos, ainda que ofertados por pessoas jurídicas;'),
        bul('Não foram considerados preços de fornecedores que não emitem Nota Fiscal Eletrônica (NF-e);'),
        bul('Nas plataformas Mercado Livre e Amazon, foram considerados exclusivamente preços ofertados por empresas (PJ) com emissão comprovada de NF-e, conforme art. 5º, III da IN 65/2021.'), sp(),
        sec('5. VALIDADE E OBSERVAÇÕES'),
        para(
            'Os preços coletados refletem os valores praticados no mercado na data de realização da pesquisa e são '
            'válidos para fins de estimativa do valor da contratação, nos termos do art. 23 da Lei nº 14.133/2021. '
            'Eventuais variações de preço ocorridas após a data de coleta deverão ser objeto de nova pesquisa '
            'antes da elaboração do processo licitatório.'
        ),
        para('A presente pesquisa não substitui o procedimento licitatório e tem caráter meramente estimativo.'), sp(),
        sec('6. RESPONSABILIDADE'),
        para(
            f'Os dados desta pesquisa foram obtidos com auxílio de ferramenta de inteligência artificial e devem ser '
            f'revisados e validados pelo servidor responsável pela instrução do processo administrativo de contratação '
            f'antes de sua utilização oficial. Pesquisa realizada em: <b>{date_str}</b>.'
        ), sp(16),
        # Signature table
        Table(
            [[
                Paragraph('_' * 50, body_sty),
                Paragraph('', body_sty),
                Paragraph('_' * 50, body_sty),
            ],[
                Paragraph('Servidor Responsável pela Pesquisa<br/>Matrícula / SIAPE: _______________', body_sty),
                Paragraph('', body_sty),
                Paragraph('Chefe Imediato / Autoridade Competente<br/>Matrícula / SIAPE: _______________', body_sty),
            ]],
            colWidths=[90*mm, 30*mm, 90*mm],
        ),
    ]

    # ── page template ──
    def on_page(canvas, doc):
        canvas.saveState()
        w, h = landscape(A4)
        # green header
        canvas.setFillColor(GREEN)
        canvas.rect(0, h - 26*mm, w, 26*mm, fill=1, stroke=0)
        canvas.setFillColor(WHITE)
        canvas.setFont('Helvetica-Bold', 13)
        canvas.drawCentredString(w/2, h - 11*mm, 'INSTITUTO FEDERAL DE ALAGOAS — IFAL')
        canvas.setFont('Helvetica', 9)
        if doc.page == 1 or (hasattr(doc, '_meta_start') and doc.page < doc._meta_start):
            canvas.drawCentredString(w/2, h - 17*mm, 'PESQUISA DE PREÇOS DE MERCADO')
            canvas.setFont('Helvetica', 7.5)
            canvas.drawCentredString(w/2, h - 22*mm,
                f'Lei nº 14.133/2021  |  IN SEGES/ME nº 65/2021  |  Data: {date_str}')
        else:
            canvas.drawCentredString(w/2, h - 15*mm, 'METODOLOGIA E PARÂMETROS DA PESQUISA DE PREÇOS')
            canvas.setFont('Helvetica', 8)
            canvas.drawCentredString(w/2, h - 21*mm, f'Instituto Federal de Alagoas — IFAL  |  {date_str}')
        # footer
        canvas.setFillColor(colors.HexColor('#888888'))
        canvas.setFont('Helvetica', 7)
        canvas.drawString(LM, 8*mm, 'IFAL — Pesquisa de Preços')
        canvas.drawRightString(w - RM, 8*mm, f'Página {doc.page}')
        canvas.restoreState()

    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=LM, rightMargin=RM,
        topMargin=TM, bottomMargin=BM,
    )

    story = [tbl, PageBreak()] + methodology
    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)

    buf.seek(0)
    return send_file(
        buf,
        mimetype='application/pdf',
        as_attachment=True,
        download_name=f'pesquisa-precos-ifal-{datetime.now().strftime("%d-%m-%Y")}.pdf'
    )


# ─── Rota principal ──────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template_string(HTML)


# ─── Inicialização ───────────────────────────────────────────────────────────

if __name__ == '__main__':
    key = os.environ.get('ANTHROPIC_API_KEY', '')
    if not key:
        print()
        print('  ⚠️  ATENÇÃO: Variável ANTHROPIC_API_KEY não encontrada!')
        print('  Configure-a antes de usar a pesquisa de preços:')
        print('    Windows:   set ANTHROPIC_API_KEY=sk-ant-...')
        print('    Mac/Linux: export ANTHROPIC_API_KEY=sk-ant-...')
        print()
    else:
        print(f'  ✓  API Key configurada.')
    print('  🌐  Acesse o sistema em: http://localhost:5000')
    print()
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
