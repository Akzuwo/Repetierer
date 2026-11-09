// Scroll position drives both entry and exit; there is no idle animation loop.
let instance;
function initScrollEffects() {
 if(instance)return instance;
 const preference=matchMedia('(prefers-reduced-motion: reduce)'),lifecycle=new AbortController(),signal=lifecycle.signal;
 const items=new Map(),scrolls=new WeakMap();let frame=0,disposed=false;
 const selector=['#class-list > div > button','.attendance-person','.participation-student','.stats-class-card:not(.stats-class-detail)','.stats-class-heading','.stats-metric','.stats-detail','.stats-timing-note','.settings-scroll-area > *','.excel-editor-item','.assessment-round-card','.assessment-student-card','.teacher-criterion','.ai-message:not(.ai-user)','.settings-guide > *','.backup-item','.log-entry','.person-item','.probability-item','.assessment-student-row','.teacher-criterion-row','.teacher-comments > *','#update-news-body > *','#class-import-preview > *','#participation-summary-list > *','tbody > tr','[data-scroll-reveal]'].join(',');
 const rootFor=element=>{for(let p=element.parentElement;p;p=p.parentElement)if(/auto|scroll/.test(getComputedStyle(p).overflowY))return p;return null;};
 const restore=(element,state)=>{element.style.translate=state.originalTranslate;element.style.opacity=state.originalOpacity;state.shift=0;};
 const observer=new IntersectionObserver(entries=>{
  for(const entry of entries){const state=items.get(entry.target);if(!state)continue;
   if(!entry.isIntersecting){state.armed=true;continue;}
   if(state.armed&&scrolls.has(state.root)&&!preference.matches&&!document.hidden)entry.target.dataset.scrollEntries=String(Number(entry.target.dataset.scrollEntries||0)+1);
   state.armed=false;
  }
 },{threshold:.06});
 function paint(){
  frame=0;if(disposed)return;
  const roots=new Map(),updates=[];
  for(const [element,state] of items){
   if(!element.isConnected){observer.unobserve(element);items.delete(element);continue;}
   if(preference.matches||document.hidden||!scrolls.has(state.root)||!element.checkVisibility({checkVisibilityCSS:true})){updates.push([element,state,0,state.baseOpacity]);continue;}
   let root=roots.get(state.root);if(!root){root=state.root.getBoundingClientRect();roots.set(state.root,root);}
   const box=element.getBoundingClientRect(),center=(box.top+box.bottom)/2-state.shift,zone=Math.min(86,root.height*.24);
   let edge=0;
   if(state.root.scrollTop>1&&center<root.top+zone)edge=-Math.min(1,(root.top+zone-center)/zone);
   else if(state.root.scrollTop<state.root.scrollHeight-state.root.clientHeight-1&&center>root.bottom-zone)edge=Math.min(1,(center-root.bottom+zone)/zone);
   const amount=Math.abs(edge),shift=Math.sign(edge)*76*Math.pow(amount,1.2),opacity=state.baseOpacity*Math.pow(1-amount,1.5);
   updates.push([element,state,shift,opacity]);
  }
  for(const [element,state,shift,opacity] of updates){state.shift=shift;element.style.translate=shift?`0 ${shift}px`:state.originalTranslate;element.style.opacity=shift?String(opacity):state.originalOpacity;}
  window.repetiererKristall?.clock.animateFor(90);
 }
 function refresh(){
  for(const [element,state] of items)if(!element.isConnected){observer.unobserve(element);items.delete(element);}
  for(const element of document.querySelectorAll(selector)){
   if(items.has(element))continue;const root=rootFor(element);if(!root||element.parentElement.closest(selector))continue;
   const state={root,armed:false,shift:0,baseOpacity:element.matches(':disabled,.disabled')?Number(getComputedStyle(element).opacity):1,originalOpacity:element.style.opacity,originalTranslate:element.style.translate};
   items.set(element,state);observer.observe(element);
  }
  schedule();
 }
 function schedule(){if(!frame&&!disposed)frame=requestAnimationFrame(paint);}
 const mutations=new MutationObserver(records=>{if(records.some(r=>r.type==='attributes'||[...r.addedNodes,...r.removedNodes].some(n=>n.nodeType===1)))refresh();});
 mutations.observe(document.getElementById('container'),{subtree:true,childList:true,attributes:true,attributeFilter:['class','hidden']});
 document.addEventListener('scroll',event=>{const root=event.target===document?document.scrollingElement:event.target;if(!root||typeof root.scrollTop!=='number')return;const previous=scrolls.get(root);scrolls.set(root,{top:root.scrollTop,direction:root.scrollTop<(previous?.top||0)?-1:1});schedule();},{capture:true,passive:true,signal});
 const finish=()=>{for(const [element,state]of items)restore(element,state);};
 preference.addEventListener('change',()=>{if(preference.matches)finish();else schedule();},{signal});
 document.addEventListener('visibilitychange',()=>{if(document.hidden)finish();else schedule();},{signal});
 window.addEventListener('resize',refresh,{signal});refresh();
 instance={refresh,schedule,dispose(){disposed=true;lifecycle.abort();observer.disconnect();mutations.disconnect();cancelAnimationFrame(frame);finish();items.clear();instance=null;}};
 return instance;
}
module.exports={initScrollEffects};
