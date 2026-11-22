(function(){
 'use strict';
 const key='volunteer-workbench-theme';
 const system=window.matchMedia('(prefers-color-scheme: dark)');
 let preference=null;
 try{const saved=localStorage.getItem(key);if(saved==='dark'||saved==='light')preference=saved;}catch{}
 function apply(){
  const dark=(preference|| (system.matches?'dark':'light'))==='dark';
  document.documentElement.dataset.theme=dark?'dark':'light';
  const button=document.getElementById('theme-toggle');
  if(button){
   button.setAttribute('aria-pressed',String(dark));
   button.setAttribute('aria-label',dark?'切换为浅色模式':'切换为深色模式');
   button.title=dark?'切换为浅色模式':'切换为深色模式';
   button.querySelector('span').textContent=dark?'浅色':'深色';
  }
 }
 apply();
 system.addEventListener('change',()=>{if(!preference)apply();});
 window.addEventListener('storage',event=>{
  if(event.key!==key&&event.key!==null)return;
  preference=event.newValue==='dark'||event.newValue==='light'?event.newValue:null;apply();
 });
 document.addEventListener('DOMContentLoaded',()=>{
  apply();
  document.getElementById('theme-toggle').addEventListener('click',()=>{
   preference=document.documentElement.dataset.theme==='dark'?'light':'dark';
   try{localStorage.setItem(key,preference);}catch{}
   apply();
  });
 });
})();
