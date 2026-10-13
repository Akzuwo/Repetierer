(() => {
  'use strict';
  const roundId = new URL(location.href).searchParams.get('round');
  const form = document.getElementById('survey');
  const status = document.getElementById('status');
  const submitStatus = document.getElementById('submit-status');
  const submit = document.getElementById('submit');
  const retry = document.getElementById('retry');
  let round;
  let browserToken;
  const key = 'repetierer-round-' + roundId;
  const messages = {duplicate_response:'Für diesen Namen oder Browser wurde bereits eine Antwort abgegeben. Wende dich bei einem Fehler an deine Lehrperson.', round_closed:'Diese Runde ist geschlossen. Es sind keine weiteren Abgaben möglich.', round_closed_or_full:'Diese Runde ist geschlossen oder vollständig. Wende dich an deine Lehrperson.', round_not_found:'Diese Runde wurde nicht gefunden oder bereits gelöscht.', invalid_email:'Bitte prüfe deine E-Mail-Adresse.', invalid_answers:'Bitte beantworte alle Aussagen.'};
  async function api(path, options) {
    const response = await fetch('./api/public/rounds/' + roundId + path, {...options, signal:AbortSignal.timeout(20000)});
    const result = await response.json();
    if (!response.ok) throw new Error(messages[result.error] || 'Die Anfrage konnte nicht verarbeitet werden. Bitte versuche es erneut.');
    return result;
  }
  function notice(element, message, error = false) {element.textContent=message; element.dataset.error=String(error);}
  async function load() {
    retry.hidden = true;
    if (!/^[a-f0-9-]{36}$/.test(roundId || '')) { notice(status,'Öffne den vollständigen Umfragelink deiner Lehrperson.'); return; }
    try {
      ({round} = await api(''));
      document.getElementById('round-title').textContent = round.title;
      if (round.closed) {notice(status,messages.round_closed); return;}
      let saved;
      try {saved=JSON.parse(localStorage.getItem(key) || 'null');} catch {}
      if (saved?.expiresAt <= Date.now()) {try{localStorage.removeItem(key);}catch{} saved=null;}
      if (saved?.submitted) {notice(status,'Deine Selbstbeurteilung wurde bereits abgegeben. Vielen Dank.'); return;}
      browserToken = saved?.token || Array.from(crypto.getRandomValues(new Uint8Array(24)), b=>b.toString(16).padStart(2,'0')).join('');
      try {
        for (let i=localStorage.length-1;i>=0;i--) {
          const item=localStorage.key(i);
          if (item?.startsWith('repetierer-round-')) {try {if(JSON.parse(localStorage.getItem(item)).expiresAt<=Date.now()) localStorage.removeItem(item);} catch {localStorage.removeItem(item);}}
        }
        localStorage.setItem(key,JSON.stringify({token:browserToken, expiresAt:Date.parse(round.expiresAt)}));
      } catch {}
      const criteria = document.getElementById('criteria');
      criteria.replaceChildren();
      for (const criterion of round.criteria) {
        const fieldset=document.createElement('fieldset');
        const legend=document.createElement('legend'); legend.textContent=criterion.label; fieldset.append(legend);
        const choices=document.createElement('div'); choices.className='choices';
        criterion.scaleLabels.forEach((label,index)=>{
          const option=document.createElement('label'); option.className='choice';
          const input=document.createElement('input'); input.type='radio'; input.name=criterion.id; input.value=String(index+1); input.required=true;
          const span=document.createElement('span'); const number=document.createElement('b'); number.textContent=String(index+1); span.append(number,document.createTextNode(label));
          option.append(input,span); choices.append(option);
        });
        fieldset.append(choices); criteria.append(fieldset);
      }
      notice(status,'Abgabe bis ' + new Intl.DateTimeFormat('de-CH',{dateStyle:'long',timeStyle:'short',timeZone:'Europe/Zurich'}).format(new Date(round.deadline)) + '.');
      form.hidden=false;
    } catch {notice(status,'Die Umfrage konnte nicht geladen werden. Prüfe deine Verbindung und versuche es erneut.',true); retry.hidden=false;}
  }
  retry.addEventListener('click',load);
  form.addEventListener('submit',async event=>{
    event.preventDefault();
    if (!form.reportValidity() || submit.disabled) return;
    const data=new FormData(form);
    submit.disabled=true; submit.textContent='Wird abgegeben …'; notice(submitStatus,'');
    try {
      await api('/responses',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name:data.get('name'),email:data.get('email'),browserToken,answers:Object.fromEntries(round.criteria.map(c=>[c.id,Number(data.get(c.id))])),additionalQuality:data.get('additionalQuality'),comment:data.get('comment')})});
      try {localStorage.setItem(key,JSON.stringify({token:browserToken,submitted:true,expiresAt:Date.parse(round.expiresAt)}));} catch {}
      form.hidden=true; notice(status,'Vielen Dank. Deine Selbstbeurteilung wurde abgegeben.'); status.tabIndex=-1; status.focus();
    } catch(error) {notice(submitStatus,error.message,true);}
    finally {submit.disabled=false;submit.textContent='Selbstbeurteilung abgeben';}
  });
  void load();
})();
