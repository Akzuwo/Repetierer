const motion=require('./uiMotion');
let modal,active,closing=false;
const queue=[];
function initQuestions(){
 if(modal)return;
 modal=document.createElement('div');modal.id='question-modal';modal.className='modal';
 modal.setAttribute('role','dialog');modal.setAttribute('aria-modal','true');modal.setAttribute('aria-labelledby','question-title');modal.setAttribute('aria-describedby','question-content');modal.setAttribute('aria-hidden','true');
 modal.innerHTML='<div class="modal-content question-content"><h2 id="question-title"></h2><p id="question-content"></p><div class="modal-buttons"><button type="button" class="btn-2" id="question-button1"></button><button type="button" class="btn-2" id="question-button2"></button></div></div>';
 document.getElementById('container').append(modal);
 modal.querySelector('#question-button1').addEventListener('click',event=>{if(event.detail<2)settle('button1','button');});
 modal.querySelector('#question-button2').addEventListener('click',event=>{if(event.detail<2)settle('button2','button');});
 modal.addEventListener('click',event=>{if(event.target===modal){event.stopImmediatePropagation();settle('cancel','backdrop');}},true);
 document.addEventListener('keydown',event=>{
  if(!active&&!closing)return;
  if(event.key==='Escape'){event.preventDefault();event.stopImmediatePropagation();settle('cancel','escape');}
  if(event.key==='Tab'){
   event.preventDefault();event.stopImmediatePropagation();
   const items=[...modal.querySelectorAll('button,[tabindex="0"]')];
   const index=items.indexOf(document.activeElement);items[(index+(event.shiftKey?-1:1)+items.length)%items.length]?.focus();
  }
 },true);
 window.askQuestion=askQuestion;
}
function askQuestion(options){
 initQuestions();
 if(!options||typeof options.title!=='string'||!options.title.trim()||options.title.length>160||typeof options.content!=='string'||options.content.length>8000||!Array.isArray(options.buttons)||options.buttons.length!==2||options.buttons.some(text=>typeof text!=='string'||!text.trim()||text.length>90)||![1,2,null,undefined].includes(options.highlight)||![1,2,undefined].includes(options.focus)||![1,2,null,undefined].includes(options.destructive))return Promise.reject(new TypeError('Ungültige Frageparameter.'));
 return new Promise(resolve=>{queue.push({options:{highlight:1,focus:1,...options,buttons:[...options.buttons]},resolve});next();});
}
function next(){
 if(active||closing||!queue.length)return;
 active=queue.shift();active.previousFocus=document.activeElement;
 window.repetiererSelects?.closeAll();
 active.background=[...document.getElementById('container').children].filter(element=>element!==modal).map(element=>({element,inert:element.inert}));
 active.background.forEach(({element})=>{element.inert=true;});
 const {options}=active;
 modal.querySelector('h2').textContent=options.title;
 const text=modal.querySelector('p');text.textContent=options.content;
 for(let index=1;index<=2;index++){
  const button=modal.querySelector(`#question-button${index}`);button.textContent=options.buttons[index-1];button.className='btn-2 question-action';
  if(button.hasAttribute('data-glass'))button.classList.add('crystal-glass');
  button.classList.toggle('question-highlight',options.highlight===index);button.classList.toggle('question-danger',options.destructive===index);
 }
 motion.openDialog(modal);window.repetiererKristall?.refresh();
 text.tabIndex=text.scrollHeight>text.clientHeight?0:-1;
 modal.querySelector(`#question-button${options.focus}`).focus({preventScroll:true});
}
function settle(action,reason){
 if(!active||closing)return;
 closing=true;const request=active;active=null;
 motion.closeDialog(modal,()=>{
  request.background.forEach(({element,inert})=>{if(element.isConnected)element.inert=inert;});
  if(request.previousFocus?.isConnected&&!request.previousFocus.closest('[inert]'))request.previousFocus.focus({preventScroll:true});
  closing=false;request.resolve({action,reason});next();
 });
}
module.exports={initQuestions,askQuestion};
