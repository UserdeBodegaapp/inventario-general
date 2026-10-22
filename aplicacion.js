"use strict";
const CONFIG = { dataUrl: "asignaciones.json", salt: "INV-GENERAL-2026-V1", minLength: 3 };
const $ = (id) => document.getElementById(id);
const form = $("formConsulta"), input = $("eNumber"), btn = $("btnConsultar"), mensaje = $("mensaje"), resultado = $("resultado"), limpiar = $("limpiar");
let asignaciones = null;
function normalizar(valor){ return String(valor || "").trim().toUpperCase().replace(/[^A-Z0-9]/g, ""); }
async function sha256(texto){ const bytes = new TextEncoder().encode(texto); const hash = await crypto.subtle.digest("SHA-256", bytes); return [...new Uint8Array(hash)].map(b=>b.toString(16).padStart(2,"0")).join(""); }
async function cargarDatos(){ if(asignaciones) return asignaciones; const r=await fetch(CONFIG.dataUrl,{cache:"no-store"}); if(!r.ok) throw new Error("No fue posible cargar asignaciones.json"); const data=await r.json(); if(!Array.isArray(data)) throw new Error("Formato de datos inválido"); asignaciones=data; return data; }
function mostrarMensaje(texto,tipo="error"){ mensaje.textContent=texto; mensaje.className="message "+(tipo==="success"?"success":""); }
function setLoading(activo){ form.classList.toggle("is-loading",activo); btn.disabled=activo; btn.querySelector("span").textContent=activo?"CONSULTANDO":"CONSULTAR ASIGNACIÓN"; }
function ocultarResultado(){ resultado.hidden=true; }
function pintar(reg){ $("resPareja").textContent=reg.pareja||"-"; $("resMesa").textContent=reg.mesa||"-"; $("resListado").textContent=reg.listado||"-"; $("resUbicacion").textContent=reg.ubicacion||"-"; resultado.hidden=false; mostrarMensaje("Consulta completada.","success"); setTimeout(()=>resultado.scrollIntoView({behavior:"smooth",block:"nearest"}),80); }
form.addEventListener("submit",async(e)=>{ e.preventDefault(); ocultarResultado(); mostrarMensaje(""); const codigo=normalizar(input.value); if(codigo.length<CONFIG.minLength){ mostrarMensaje("Ingresa un E-Number válido para continuar."); input.focus(); return; } setLoading(true); try{ const hash=await sha256(CONFIG.salt+"|"+codigo); const datos=await cargarDatos(); const reg=datos.find(x=>x.id===hash); if(!reg){ mostrarMensaje("No encontramos una asignación. Verifica el E-Number o consulta con el equipo de inventario."); return; } pintar(reg); }catch(err){ console.error(err); mostrarMensaje("No fue posible realizar la consulta. Revisa tu conexión e inténtalo nuevamente."); }finally{ setLoading(false); } });
input.addEventListener("input",()=>{ input.value=input.value.toUpperCase(); limpiar.hidden=!input.value; mostrarMensaje(""); ocultarResultado(); });
limpiar.addEventListener("click",()=>{ input.value=""; limpiar.hidden=true; mostrarMensaje(""); ocultarResultado(); input.focus(); });
$("nuevaConsulta").addEventListener("click",()=>{ input.value=""; limpiar.hidden=true; mostrarMensaje(""); ocultarResultado(); input.focus(); window.scrollTo({top:0,behavior:"smooth"}); });
