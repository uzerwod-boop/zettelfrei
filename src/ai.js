/* Zettelfrei – KI-Anbindung.
 * Anbieter:
 *   server : eigener Node-Server (server.js) mit ANTHROPIC_API_KEY, Schlüssel bleibt auf dem Server
 *   claude : Claude API direkt aus dem Browser mit eigenem Schlüssel
 *   local  : OpenAI-kompatibler Endpunkt, z. B. Ollama oder LM Studio auf dem eigenen Rechner
 *   artifact: läuft die Seite als Claude-Artefakt, nutzt sie Claude direkt
 * Alle Anbieter liefern über AI.json() dasselbe: das geparste JSON der Antwort.
 */
var AI = (function(){
"use strict";
var KEY="zf-ai-settings";
var DEFAULT_MODEL="claude-sonnet-5-5";
var cfg=load(), artifactSample=null, server=null, listeners=[];

function load(){ try{ var c=JSON.parse(localStorage.getItem(KEY)||"null"); if(c) return c; }catch(e){} return {provider:"auto",key:"",model:"",url:"http://localhost:11434/v1"}; }
function save(c){ cfg=Object.assign({},cfg,c); try{ localStorage.setItem(KEY,JSON.stringify(cfg)); }catch(e){} notify(); }
function notify(){ listeners.forEach(function(f){ try{ f(status()); }catch(e){} }); }

function init(){
  var tasks=[];
  if(window.claude && window.claude.use){
    tasks.push(window.claude.use("sample").then(function(s){ artifactSample=s; },function(){}));
  }
  if(location.protocol.indexOf("http")===0){
    tasks.push(fetch("api/health",{cache:"no-store"}).then(function(r){ return r.ok?r.json():null; }).then(function(j){ server=j&&j.ok?j:null; },function(){ server=null; }));
  }
  return Promise.all(tasks).then(notify);
}

function active(){
  var p=cfg.provider||"auto";
  if(p==="off") return null;
  if(p==="auto"){ if(artifactSample) return "artifact"; if(server) return "server"; if(cfg.key) return "claude"; return null; }
  if(p==="server") return server?"server":null;
  if(p==="claude") return cfg.key?"claude":null;
  if(p==="local") return cfg.url?"local":null;
  return null;
}
function status(){
  var a=active();
  var names={artifact:"Claude",server:"Server"+(server&&server.model?" · "+server.model:""),claude:"Claude API · "+(cfg.model||DEFAULT_MODEL),local:"Lokal · "+(cfg.model||"Modell")};
  return {on:!!a, provider:a, label:a?names[a]:(cfg.provider==="off"?"KI aus":"KI nicht eingerichtet"), cfg:cfg, server:!!server};
}

/* ---------- Hilfen ---------- */
function err(code,message){ return {code:code,message:message}; }
function parseJson(text){
  var t=String(text||"").trim();
  try{ return JSON.parse(t); }catch(e){}
  var m=t.match(/```(?:json)?\s*([\s\S]*?)```/); if(m){ try{ return JSON.parse(m[1]); }catch(e){} }
  var a=Math.min.apply(null,[t.indexOf("{"),t.indexOf("[")].filter(function(i){ return i>=0; }).concat([Infinity]));
  var b=Math.max(t.lastIndexOf("}"),t.lastIndexOf("]"));
  if(a!==Infinity && b>a){ try{ return JSON.parse(t.slice(a,b+1)); }catch(e){} }
  var x=err("invalid_json","Antwort enthielt kein gültiges JSON"); x.text=t; throw x;
}
function toImage(blob){ // verkleinert auf max. 1568 px, liefert {mediaType, data(base64)}
  return new Promise(function(resolve,reject){
    var url=URL.createObjectURL(blob), img=new Image();
    img.onload=function(){
      var s=Math.min(1,1568/Math.max(img.width,img.height)), c=document.createElement("canvas");
      c.width=Math.round(img.width*s); c.height=Math.round(img.height*s);
      var g=c.getContext("2d"); g.fillStyle="#fff"; g.fillRect(0,0,c.width,c.height); g.drawImage(img,0,0,c.width,c.height);
      URL.revokeObjectURL(url);
      var d=c.toDataURL("image/jpeg",0.88); resolve({mediaType:"image/jpeg",data:d.split(",")[1]});
    };
    img.onerror=function(){ URL.revokeObjectURL(url); reject(err("image_rejected","Bild nicht lesbar")); };
    img.src=url;
  });
}
function httpError(status,body){
  if(status===401||status===403) return err("auth","Schlüssel ungültig oder ohne Berechtigung");
  if(status===429) return err("rate_limited","Zu viele Anfragen");
  if(status===413) return err("prompt_too_large","Anfrage zu groß");
  if(status===400 && /too long|too large|context/i.test(body)) return err("prompt_too_large","Anfrage zu groß");
  var e=err("http","Fehler "+status); e.detail=body; return e;
}
function post(url,headers,body,signal){
  return fetch(url,{method:"POST",headers:headers,body:JSON.stringify(body),signal:signal}).then(function(r){
    if(!r.ok) return r.text().then(function(t){ throw httpError(r.status,t); });
    return r.json();
  },function(e){ if(e&&e.name==="AbortError") throw err("cancelled","abgebrochen"); throw err("network","Keine Verbindung zum KI-Anbieter"); });
}

/* ---------- Anfragen ---------- */
var JSON_HINT="\n\nAntworte ausschließlich mit gültigem JSON, ohne Text davor oder danach.";
function anthropicBody(prompt,images,model){
  var content=images.map(function(im){ return {type:"image",source:{type:"base64",media_type:im.mediaType,data:im.data}}; });
  content.push({type:"text",text:prompt+JSON_HINT});
  return {model:model||DEFAULT_MODEL,max_tokens:4096,messages:[{role:"user",content:content}]};
}
function anthropicText(res){ return (res.content||[]).filter(function(b){ return b.type==="text"; }).map(function(b){ return b.text; }).join(""); }

function json(prompt,opts){
  opts=opts||{};
  var p=active();
  if(!p) return Promise.reject(err("no_provider","Keine KI eingerichtet"));
  if(p==="artifact"){
    var o={cache:false}; if(opts.signal) o.signal=opts.signal; if(opts.images&&opts.images.length) o.images=opts.images;
    return artifactSample.json(prompt,o);
  }
  var imgs=(opts.images||[]);
  return Promise.all(imgs.map(toImage)).then(function(images){
    if(p==="server") return post("api/messages",{"content-type":"application/json"},anthropicBody(prompt,images,null),opts.signal).then(anthropicText);
    if(p==="claude") return post("https://api.anthropic.com/v1/messages",{
        "content-type":"application/json","x-api-key":cfg.key,"anthropic-version":"2023-06-01","anthropic-dangerous-direct-browser-access":"true"
      },anthropicBody(prompt,images,cfg.model),opts.signal).then(anthropicText);
    // local: OpenAI-kompatibel
    var content=[{type:"text",text:prompt+JSON_HINT}].concat(images.map(function(im){ return {type:"image_url",image_url:{url:"data:"+im.mediaType+";base64,"+im.data}}; }));
    var headers={"content-type":"application/json"}; if(cfg.key) headers.authorization="Bearer "+cfg.key;
    var base=String(cfg.url||"").replace(/\/+$/,"");
    return post(base+"/chat/completions",headers,{model:cfg.model||"llama3.1",temperature:0,messages:[{role:"user",content:images.length?content:prompt+JSON_HINT}]},opts.signal)
      .then(function(res){ return res.choices&&res.choices[0]&&res.choices[0].message?res.choices[0].message.content:""; });
  }).then(function(text){ if(typeof text!=="string"||!text.trim()) throw err("empty_completion","Leere Antwort"); return parseJson(text); });
}

function test(){
  return json('Antworte mit {"ok": true}.').then(function(r){ return !!(r&&r.ok); });
}

return {init:init, json:json, test:test, status:status, save:save, active:active, onChange:function(f){ listeners.push(f); }, parseJson:parseJson, DEFAULT_MODEL:DEFAULT_MODEL};
})();
