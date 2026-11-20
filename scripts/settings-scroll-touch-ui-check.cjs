const assert = require('node:assert/strict');

module.exports = async function ({win, evaluate, waitFor, shot, pause}) {
  const cdp = (name, args) => win.webContents.debugger.sendCommand(name, args);
  const mouse = async selector => {
    const point = await evaluate(`(() => {const r=document.querySelector(${JSON.stringify(selector)}).getBoundingClientRect();return {x:Math.round(r.x+r.width/2),y:Math.round(r.y+r.height/2)};})()`);
    win.webContents.sendInputEvent({type:'mouseDown',button:'left',clickCount:1,...point});
    win.webContents.sendInputEvent({type:'mouseUp',button:'left',clickCount:1,...point});
    await pause(60);
  };
  const settled = () => waitFor(`!document.getElementById('content').dataset.pageMotion && !repetiererKristall.clock.frame`);
  win.show(); win.focus();
  await waitFor(`document.visibilityState==='visible' && repetiererKristall.renderer?.ready`);
  await mouse('#settings-btn'); await settled();
  await shot('settings-scroll-initial');

  // Capture the actual shader uniforms submitted at each draw. After all RAF
  // callbacks, they must still match the DOM painted in that very frame.
  await evaluate(`(() => {
    const renderer=repetiererKristall.renderer, gl=renderer.gl;
    window.scrollLensAudit={draws:new Map(),activeElement:null};
    const getRect=Element.prototype.getBoundingClientRect;
    Element.prototype.getBoundingClientRect=function(){const rect=getRect.call(this);if(window.scrollLensAudit.recording && this.hasAttribute('data-glass'))scrollLensAudit.activeElement=this;return rect;};
    const uniform=gl.uniform4f.bind(gl);
    gl.uniform4f=function(location,x,y,w,h){if(location===renderer.uniforms.rect && scrollLensAudit.recording)scrollLensAudit.draws.set(scrollLensAudit.activeElement,{x,y,w,h});return uniform(location,x,y,w,h);};
    const draw=renderer.draw.bind(renderer);
    renderer.draw=function(){scrollLensAudit.recording=true;try{return draw();}finally{scrollLensAudit.recording=false;}};
  })()`);
  async function scrollFrames(panel) {
    return evaluate(`(async () => {
      const root=document.querySelector('#settings-view .settings-scroll-area'), mismatches=[];
      const elements=[...document.querySelectorAll('#settings-${panel}-fields :is(.range-control,.liquid-switch,.switch-thumb,.liquid-select-trigger)')];
      let moved=0;
      for(const top of [20,45,80,120,165,210,260,310,270,220,175,125,75,25,0]) {
        root.scrollTop=top;
        // Same native event ordering as a scroll: canvas listener, then effects.
        root.dispatchEvent(new Event('scroll',{bubbles:false}));
        await new Promise(resolve=>requestAnimationFrame(resolve));
        for(const element of elements) {
          const r=element.getBoundingClientRect(), drawn=scrollLensAudit.draws.get(element), clip=root.getBoundingClientRect();
          if(!element.checkVisibility({checkOpacity:true}) || r.bottom<=clip.top || r.top>=clip.bottom || !drawn)continue;
          const delta=Math.max(Math.abs(drawn.x-r.x),Math.abs(drawn.y-r.y),Math.abs(drawn.w-r.width),Math.abs(drawn.h-r.height));
          if(delta>.1)mismatches.push({top,control:element.className,delta});
          if(element.closest('label')?.style.translate)moved++;
        }
      }
      return {mismatches,moved};
    })()`);
  }
  for (const panel of ['global','class']) {
    if (panel==='class') {await mouse('#settings-class-tab'); await settled();}
    const frames=await scrollFrames(panel);
    assert.ok(frames.moved>0,'The scroll effect ran for '+panel);
    assert.deepEqual(frames.mismatches,[],panel+': shader and HTML stay together during entry, exit and reversal');
    await shot('settings-scroll-'+panel);
  }

  // Native wheel moves the class picker with its panel; the fixed actions and
  // navigation stay in place. Opening a menu then scrolling closes it.
  await evaluate(`document.querySelector('#settings-view .settings-scroll-area').scrollTop=0`);
  await pause(120);
  const pickerBefore = await evaluate(`document.getElementById('settings-class-picker').getBoundingClientRect().top`);
  const chrome = () => evaluate(`['settings-heading','settings-back-btn','settings-navdock'].map(id=>{const r=document.getElementById(id).getBoundingClientRect();return [r.x,r.y,r.width,r.height];})`);
  const fixedBefore=await chrome();
  await mouse('#settings-class-picker .liquid-select-trigger');
  await waitFor(`document.querySelector('#settings-class-picker .liquid-select-trigger').getAttribute('aria-expanded')==='true'`);
  await cdp('Input.dispatchMouseEvent',{type:'mouseWheel',x:850,y:430,deltaY:110,deltaX:0});
  await waitFor(`document.querySelector('#settings-view .settings-scroll-area').scrollTop>40`);
  await pause(120);
  assert.ok(await evaluate(`document.getElementById('settings-class-picker').getBoundingClientRect().top<${pickerBefore}-40`),'Class picker scrolls with its panel');
  assert.deepEqual(await chrome(),fixedBefore,'Heading, back action and navigation stay fixed');
  assert.equal(await evaluate(`document.querySelector('#settings-class-picker .liquid-select-trigger').getAttribute('aria-expanded')`),'false');
  await evaluate(`document.querySelector('#settings-view .settings-scroll-area').scrollTop=0`); await pause(100);
  await mouse('#settings-class-picker .liquid-select-trigger');
  await waitFor(`document.querySelector('#settings-class-picker .liquid-select-trigger').getAttribute('aria-expanded')==='true'`);
  win.webContents.sendInputEvent({type:'keyDown',keyCode:'Escape'});
  win.webContents.sendInputEvent({type:'keyUp',keyCode:'Escape'});
  await waitFor(`document.querySelector('#settings-class-picker .liquid-select-trigger').getAttribute('aria-expanded')==='false'`);
  assert.equal(await evaluate(`document.activeElement.classList.contains('liquid-select-trigger')`),true);

  // Desktop touch input, without changing the viewport into a mobile layout.
  await cdp('Emulation.setTouchEmulationEnabled',{enabled:true,maxTouchPoints:1});
  await evaluate(`window.touchAudit=[];document.addEventListener('pointerdown',e=>touchAudit.push({trusted:e.isTrusted,type:e.pointerType}),{capture:true});window.touchActivations=0;document.getElementById('settings-back-btn').addEventListener('click',()=>touchActivations++);`);
  const touch = async (selector, name) => {
    const point=await evaluate(`(() => {const r=document.querySelector(${JSON.stringify(selector)}).getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2};})()`);
    await cdp('Input.dispatchTouchEvent',{type:'touchStart',touchPoints:[{...point,id:1,radiusX:3,radiusY:3,force:1}]});
    await pause(90); if(name)await shot(name+'-pressed');
    await cdp('Input.dispatchTouchEvent',{type:'touchEnd',touchPoints:[]});
    await pause(100);
  };
  assert.equal(await evaluate(`getComputedStyle(document.getElementById('settings-global-tab')).webkitTapHighlightColor`),'rgba(0, 0, 0, 0)','Chromium tap overlay is suppressed on pills');
  await touch('#settings-back-btn','settings-touch-back');
  await waitFor(`document.getElementById('settings-view').classList.contains('update-hidden')`);
  assert.equal(await evaluate(`touchActivations`),1,'A touch activates exactly once');
  assert.ok(await evaluate(`touchAudit.some(e=>e.trusted && e.type==='touch')`),'Browser delivered real touch pointer input');
  await settled();
  await mouse('#settings-btn'); await settled();
  // Test a toggle and range on the same desktop touch path.
  await evaluate(`document.querySelector('#settings-view .settings-scroll-area').scrollTop=0`); await pause(100);
  const old=await evaluate(`document.getElementById('homework-enabled-setting').checked`);
  await touch('#homework-enabled-setting');
  assert.equal(await evaluate(`document.getElementById('homework-enabled-setting').checked`),!old);
  await touch('#homework-enabled-setting');
  const slider=await evaluate(`(() => {const r=document.getElementById('homework-joker-setting').getBoundingClientRect();return {x:r.x+r.width*.7,y:r.y+r.height/2};})()`);
  await cdp('Input.dispatchTouchEvent',{type:'touchStart',touchPoints:[{...slider,id:1}]});
  await cdp('Input.dispatchTouchEvent',{type:'touchEnd',touchPoints:[]});
  assert.ok(await evaluate(`document.getElementById('homework-joker-setting').valueAsNumber>2`),'Touch changes the native range value');
  await cdp('Emulation.setTouchEmulationEnabled',{enabled:false});

  // Reduced motion restores HTML and the lens together, even mid-scroll.
  await cdp('Emulation.setEmulatedMedia',{features:[{name:'prefers-reduced-motion',value:'reduce'}]});
  await waitFor(`matchMedia('(prefers-reduced-motion: reduce)').matches && repetiererKristall.clock.active.size===0`);
  await pause(70);
  assert.equal(await evaluate(`[...document.querySelectorAll('#settings-class-rules > *')].some(e=>!!e.style.translate)`),false);
  await shot('settings-scroll-reduced-motion');
  win.setSize(1440,900); await pause(140);
  assert.ok(await evaluate(`document.getElementById('settings-class-picker').closest('.settings-scroll-area')!==null`));
  assert.equal(await evaluate(`(() => {const r=repetiererKristall.renderer,before=r.canvas.toDataURL();r.draw();return before===r.canvas.toDataURL();})()`),true,'Reduced-motion resize leaves the current optical frame painted');
  await shot('settings-scroll-desktop-large');
  console.log('SETTINGS_SCROLL_TOUCH_UI_PASS: same-frame shader geometry, both panels, bidirectional scroll, moving class picker, fixed chrome, menu closure, trusted desktop touch, one activation, toggle/range and reduced motion');
};
