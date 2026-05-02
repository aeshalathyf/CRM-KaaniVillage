// ============================================================================
// KAANI HOTELS CRM — MAIN APP
// ============================================================================

const { createClient } = supabase;
const sb = createClient(window.SUPABASE_CONFIG.url, window.SUPABASE_CONFIG.key);

const STATE = {
  user: null,
  profile: null,
  properties: [],
  templates: [],
  tags: [],
  currentGuests: [],
  guestsTotalCount: 0,
  guestsPage: 0,
  guestsPerPage: 50,
  selectedGuestId: null,
  chartsRendered: false,
  commissionRates: [],  // channel commission rates from DB
  selectedPropertyId: null,  // null = All Properties, or a specific property ID number
};


// Fetch all rows from a query (handles Supabase's default 1000 row limit)
async function fetchAll(query){
  const pageSize = 1000;
  let allData = [];
  let page = 0;
  while(true){
    const from = page * pageSize;
    const to = from + pageSize - 1;
    const {data, error} = await query.range(from, to);
    if(error || !data || data.length === 0) break;
    allData = allData.concat(data);
    if(data.length < pageSize) break; // last page
    page++;
  }
  return allData;
}

const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const AVCOLS = ['#E1F5EE/#0F6E56','#E6F1FB/#185FA5','#FAEEDA/#633806','#FBEAF0/#72243E','#EAF3DE/#27500A','#EEEDFE/#3C3489'];

// Calculate commission and net revenue for a stay
function calcNetRevenue(source, channelType, totalRevenue){
  if(!totalRevenue || totalRevenue === 0) return 0;
  const rates = STATE.commissionRates;
  // Exact source match
  let rate = rates.find(r => r.source_name.toLowerCase() === (source||'').toLowerCase());
  // Fall back to channel type default
  if(!rate) rate = rates.find(r => r.source_name === 'DEFAULT_' + (channelType||'OTA').toUpperCase());
  const pct = rate ? parseFloat(rate.commission_pct) : 0;
  return Math.round(totalRevenue * (1 - pct / 100) * 100) / 100;
}

function getCommissionPct(source, channelType){
  const rates = STATE.commissionRates;
  let rate = rates.find(r => r.source_name.toLowerCase() === (source||'').toLowerCase());
  if(!rate) rate = rates.find(r => r.source_name === 'DEFAULT_' + (channelType||'OTA').toUpperCase());
  return rate ? parseFloat(rate.commission_pct) : 0;
}

// 6-MONTH PLAN DATA
const PLAN_MONTHS = [
  {month:'Month 1 — April / May',focus:'Foundation',badge:'be',tasks:[
    {w:'Week 1–2',title:'Clean up data',detail:'Import Ezee reports weekly. Chase missing emails at front desk. Brief team to collect email at every check-in.'},
    {w:'Week 2–3',title:'Set up review flow',detail:'Save Google Review link as WhatsApp quick reply. Send to every checkout within 24 hours.'},
    {w:'Week 3–4',title:'First upsell',detail:'Email current stayovers with excursion menu. Target guests on night 2–3.'},
    {w:'End of month',title:'Target: 90% email capture, reviews flowing',detail:''}
  ]},
  {month:'Month 2 — May',focus:'First campaign',badge:'bh',tasks:[
    {w:'Week 1',title:'Win-back campaign',detail:'Email all April OTA guests a 10% direct booking offer. Valid June–September.'},
    {w:'Week 2',title:'WhatsApp follow-up',detail:'5 days after email, send short WhatsApp to non-responders.'},
    {w:'Week 3–4',title:'Personal outreach to Destinos group',detail:'Individual emails to 20+ Spanish guests. Frame as personal invitation to return independently.'},
    {w:'End of month',title:'Target: first direct booking from campaign',detail:''}
  ]},
  {month:'Month 3 — June',focus:'Birthday engine',badge:'be',tasks:[
    {w:'Every Monday',title:'Check birthday tab',detail:'Open CRM every Monday, send birthday messages for that week. 5 minutes max.'},
    {w:'Week 2',title:'Add returning guest benefit',detail:'Free welcome drink, early check-in, or upgrade. Add to booking page.'},
    {w:'Week 3–4',title:'Upsell all June stayovers',detail:'Refine upsell message based on May results.'},
    {w:'End of month',title:'Target: birthday messages sent weekly without fail',detail:''}
  ]},
  {month:'Month 4 — July',focus:'Mid-year push',badge:'bm',tasks:[
    {w:'Week 1',title:'Q4 early bird blast',detail:'Email full list. 12–15% off October–December travel, book by August.'},
    {w:'Week 2',title:'Segment by nationality',detail:'Spanish: group travel angle. Indian: package value with transfers and meals.'},
    {w:'Week 3–4',title:'Tag your best guests',detail:'Tag VIP, Repeat, Honeymoon. Personal messages to top 10–15.'},
    {w:'End of month',title:'Target: seasonal blast sent, segmented versions tested',detail:''}
  ]},
  {month:'Month 5 — August',focus:'Deepen relationships',badge:'bh',tasks:[
    {w:'Week 1',title:'Launch post-stay sequence',detail:'Every checkout: review request day 1, win-back offer day 21. Make permanent.'},
    {w:'Week 2–3',title:'Personal messages to top guests',detail:'Individual messages (not templates) to top 10–15 tagged guests.'},
    {w:'Week 4',title:'Collect missing emails',detail:'Cross-check local Maldivian guests. Update records after every check-in.'},
    {w:'End of month',title:'Target: post-stay sequence running for every checkout',detail:''}
  ]},
  {month:'Month 6 — September',focus:'Measure and scale',badge:'bu',tasks:[
    {w:'Week 1',title:'Review your numbers',detail:'Direct booking share (target 18–20%), email capture (95%), Google reviews (doubled).'},
    {w:'Week 2',title:'Double what worked',detail:'Which campaign got replies/bookings? Do it more often. Drop what you skipped.'},
    {w:'Week 3–4',title:'Plan next 6 months',detail:'Set up newsletter for loyal guests, consider WhatsApp broadcast list.'},
    {w:'End of month',title:'Target: 2x direct bookings vs April baseline',detail:''}
  ]}
];

const CHECKLIST_DATA = {
  daily:['Send review request to yesterday\'s checkouts','Check stayovers — anyone to upsell today','Check Overview tab for any urgent actions'],
  weekly:['Export new Ezee report and import into CRM','Send birthday messages for the coming week','Tag VIP / repeat / honeymoon guests','Follow up any unanswered guest messages','Send upsell email to stayovers on night 2 or 3'],
  monthly:['Send win-back offer to OTA guests from 30–45 days ago','Run seasonal blast if next month looks soft','Review email capture — chase missing local emails','Check dashboard — which source and nationality is growing','Refresh email templates if needed']
};

function renderPlanPane(){
  const el=document.getElementById('pane-plan');
  if(!el)return;
  el.innerHTML=`
    <div class="card">
      <div class="ch"><div class="ct">6-month CRM action plan</div><span class="badge be">Apr – Sep 2026</span></div>
      <div class="cd">Your week-by-week tasks to increase direct bookings and repeat guests. Consistent execution will move direct bookings from 10% to 18–20% within 6 months.</div>
      <div class="sg">
        <div class="sm"><div class="sl">Direct target</div><div class="sv">20%</div><div class="ss">from 10%</div></div>
        <div class="sm"><div class="sl">Email capture</div><div class="sv">95%</div><div class="ss">from 79%</div></div>
        <div class="sm"><div class="sl">Repeat rate</div><div class="sv">15%</div><div class="ss">from ~5%</div></div>
        <div class="sm"><div class="sl">Reviews</div><div class="sv">2x</div><div class="ss">Google count</div></div>
      </div>
    </div>
    ${PLAN_MONTHS.map(m=>`
      <div class="card">
        <div class="ch"><div class="ct">${m.month}</div><span class="badge ${m.badge}">${m.focus}</span></div>
        ${m.tasks.map(t=>`<div style="display:flex;gap:12px;padding:9px 0;border-bottom:1px solid #E8E6E0;align-items:flex-start"><div style="font-size:11px;color:#5F5E5A;width:80px;flex-shrink:0;padding-top:2px;font-weight:500">${t.w}</div><div style="flex:1"><div style="font-size:13px;font-weight:500">${t.title}</div>${t.detail?'<div style="font-size:12px;color:#5F5E5A;margin-top:3px;line-height:1.6">'+t.detail+'</div>':''}</div></div>`).join('')}
      </div>`).join('')}`;
}

// Known placeholder/operator emails — treated as no real email
// Add more here as you discover them
const PLACEHOLDER_EMAILS = [
  'info@destinosentreazules.com',
];
function isPlaceholderEmail(e){
  if(!e)return false;
  return PLACEHOLDER_EMAILS.includes(e.toLowerCase().trim());
}


// ------------- HELPERS -------------
function avc(n){if(!n)return AVCOLS[0].split('/');let h=0;for(const c of n)h=(h*31+c.charCodeAt(0))%AVCOLS.length;return AVCOLS[h].split('/');}
function ini(n){if(!n)return'?';const p=n.replace(/^(Mr\.|Ms\.|Mrs\.|Dr\.)\s+/i,'').trim().split(/\s+/);return((p[0]?.[0]||'')+(p[p.length-1]?.[0]||'')).toUpperCase();}
function fn(n){if(!n)return'';return n.replace(/^(Mr\.|Ms\.|Mrs\.|Dr\.)\s+/i,'').trim().split(/\s+/)[0];}
function srcShort(s){if(!s)return'—';if(s.includes('DESTINOS'))return'Destinos';if(s.includes('MAKE MY TRIP'))return'MakeMyTrip';if(s==='DIRECT BOOKING')return'Direct';return s;}
function fmtDate(d){if(!d)return'—';const dt=new Date(d);if(isNaN(dt))return d;return String(dt.getDate()).padStart(2,'0')+' '+MONTHS[dt.getMonth()]+' '+dt.getFullYear();}
function calcAge(dob){if(!dob)return null;const d=new Date(dob);if(isNaN(d))return null;const today=new Date();let a=today.getFullYear()-d.getFullYear();if(today<new Date(today.getFullYear(),d.getMonth(),d.getDate()))a--;return a;}
function bdayDays(dob){if(!dob)return null;const d=new Date(dob);if(isNaN(d))return null;const today=new Date();today.setHours(0,0,0,0);let n=new Date(today.getFullYear(),d.getMonth(),d.getDate());if(n<today)n=new Date(today.getFullYear()+1,d.getMonth(),d.getDate());return Math.round((n-today)/864e5);}
function escapeHtml(s){if(!s&&s!==0)return'';return String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c]);}
function toast(msg,isError){const t=document.createElement('div');t.className='toast'+(isError?' error':'');t.textContent=msg;document.getElementById('toast-area').appendChild(t);setTimeout(()=>t.remove(),3000);}

// ------------- AUTH -------------
async function checkSession(){
  const{data:{session}}=await sb.auth.getSession();
  if(session){
    STATE.user=session.user;
    await loadProfile();
    showApp();
  }else{
    showLogin();
  }
}

async function loadProfile(){
  const{data,error}=await sb.from('user_profiles').select('*').eq('user_id',STATE.user.id).single();
  if(error||!data){
    // First time login — create profile (admin if no other admins exist)
    const{count}=await sb.from('user_profiles').select('*',{count:'exact',head:true}).eq('role','admin');
    const role=count===0?'admin':'staff';
    const{data:newProfile}=await sb.from('user_profiles').insert({user_id:STATE.user.id,full_name:STATE.user.email.split('@')[0],role}).select().single();
    STATE.profile=newProfile;
  }else{
    STATE.profile=data;
  }
}

async function doLogin(){
  const email=document.getElementById('login-email').value.trim();
  const password=document.getElementById('login-password').value;
  const msg=document.getElementById('login-msg');
  if(!email||!password){msg.className='login-msg error';msg.textContent='Please enter both email and password';return;}
  msg.className='login-msg';msg.textContent='Signing in…';
  const{data,error}=await sb.auth.signInWithPassword({email,password});
  if(error){msg.className='login-msg error';msg.textContent=error.message;return;}
  STATE.user=data.user;
  await loadProfile();
  showApp();
}

async function doLogout(){
  await sb.auth.signOut();
  STATE.user=null;STATE.profile=null;
  location.reload();
}

function showLogin(){
  document.getElementById('login-screen').style.display='flex';
  document.getElementById('app').style.display='none';
}

async function showApp(){
  document.getElementById('login-screen').style.display='none';
  document.getElementById('app').style.display='block';
  document.getElementById('user-display').textContent=
    `${STATE.profile.full_name||STATE.user.email} (${STATE.profile.role})`;

  await loadInitialData();
  setupPropertyAccess();

  // Admin tab
  const adminTab=document.getElementById('admin-tab');
  if(adminTab) adminTab.style.display=STATE.profile.role==='admin'?'inline-block':'none';

  // Activate Overview pane explicitly
  document.querySelectorAll('.pane').forEach(p=>p.classList.remove('on'));
  document.querySelectorAll('.nav .nb').forEach(b=>b.classList.remove('on'));
  const ov=document.getElementById('pane-overview');
  if(ov) ov.classList.add('on');
  const firstBtn=document.querySelector('.nav .nb');
  if(firstBtn) firstBtn.classList.add('on');

  // Render property selector — call immediately and again after render
  renderPropertySelector();
  setTimeout(renderPropertySelector, 300);
  setTimeout(renderPropertySelector, 800);

  await renderOverview();
}


function setupPropertyAccess(){
  if(STATE.profile.role==='admin'){
    STATE.accessibleProperties=STATE.properties;
    STATE.selectedPropertyId=null;
  } else if(STATE.profile.role==='manager'){
    const assigned=STATE.profile.property_ids||[];
    STATE.accessibleProperties=assigned.length>0
      ? STATE.properties.filter(p=>assigned.includes(p.id))
      : STATE.properties;
    STATE.selectedPropertyId=null;
  } else {
    const assigned=STATE.profile.property_ids||[];
    STATE.accessibleProperties=assigned.length>0
      ? STATE.properties.filter(p=>assigned.includes(p.id))
      : STATE.properties;
    STATE.selectedPropertyId=STATE.profile.default_property_id
      || (STATE.accessibleProperties[0]?.id || null);
  }
}

function renderPropertySelector(){
  const bar=document.getElementById('prop-bar');
  const pills=document.getElementById('prop-pills');
  if(!bar||!pills)return;

  // Always use properties directly — don't depend on accessibleProperties timing
  const propsToShow = STATE.accessibleProperties && STATE.accessibleProperties.length > 0
    ? STATE.accessibleProperties
    : STATE.properties;

  if(propsToShow.length === 0){
    // Properties not loaded yet — try again shortly
    setTimeout(renderPropertySelector, 200);
    return;
  }

  bar.style.display='flex';

  let html='';
  // Always show "All Properties" for admin
  const role = STATE.profile?.role || 'staff';
  if(role === 'admin' || role === 'manager'){
    html+=`<button class="pp ${STATE.selectedPropertyId===null?'on':''}" onclick="selectProperty(null)">${role==='admin'?'All Properties':'All My Properties'}</button>`;
  }
  propsToShow.forEach(p=>{
    const isOn = parseInt(STATE.selectedPropertyId)===parseInt(p.id);
    html+=`<button class="pp ${isOn?'on':''}" onclick="selectProperty(${p.id})">${escapeHtml(p.name)}</button>`;
  });

  pills.innerHTML=html;
  console.log('Property selector rendered:', propsToShow.length+1, 'pills');
}

async function selectProperty(propId){
  // Ensure consistent type — null for All Properties, integer for specific property
  STATE.selectedPropertyId = (propId === null || propId === 'null' || propId === '') ? null : parseInt(propId);
  STATE.guestsPage=0;
  STATE.chartsRendered=false;
  renderPropertySelector();

  const activePane=document.querySelector('.pane.on');
  if(!activePane){toast(STATE.selectedPropertyId?`${STATE.properties.find(p=>p.id===STATE.selectedPropertyId)?.name}`:'All Properties');return;}
  const tabId=activePane.id.replace('pane-','');
  if(tabId==='overview')renderOverview();
  else if(tabId==='dashboard')renderDashboard();
  else if(tabId==='guests'){STATE.guestsPage=0;renderGuestsPane();}
  else if(tabId==='actions')renderActionsPane();
  else if(tabId==='marketing')renderMarketingPane();
  else if(tabId==='reports')renderReportsPane();

  toast(STATE.selectedPropertyId
    ? `Viewing: ${STATE.properties.find(p=>p.id===STATE.selectedPropertyId)?.name||''}`
    : 'Viewing: All Properties');
}

async function loadInitialData(){
  const[propsRes,tagsRes,tplRes,countRes]=await Promise.all([
    sb.from('properties').select('*').order('id'),
    sb.from('tags').select('*').order('name'),
    sb.from('email_templates').select('*').order('id'),
    sb.from('guests').select('*',{count:'exact',head:true})
  ]);
  STATE.properties=propsRes.data||[];
  STATE.tags=tagsRes.data||[];
  STATE.templates=tplRes.data||[];
  STATE.guestsTotalCount=countRes.count||0;
  document.getElementById('app-subtitle').textContent=`${STATE.guestsTotalCount.toLocaleString()} guests · ${STATE.properties.filter(p=>p.active).length} properties`;
}

// ------------- TAB SWITCHING -------------

// Returns array of property IDs to filter by — NEVER empty
function getFilterPropertyIds(){
  if(STATE.selectedPropertyId !== null){
    return [parseInt(STATE.selectedPropertyId)];
  }
  const props = STATE.accessibleProperties && STATE.accessibleProperties.length > 0
    ? STATE.accessibleProperties
    : STATE.properties;
  return props.map(p=>parseInt(p.id));
}

// Check if viewing all properties
function isAllPropertiesMode(){
  return STATE.selectedPropertyId === null;
}

function sw(tab,btn){
  document.querySelectorAll('.pane').forEach(p=>p.classList.remove('on'));
  document.querySelectorAll('.nav .nb').forEach(b=>b.classList.remove('on'));
  const pane=document.getElementById('pane-'+tab);
  if(pane)pane.classList.add('on');
  btn.classList.add('on');
  if(tab==='overview')renderOverview();
  if(tab==='dashboard')renderDashboard();
  if(tab==='guests')renderGuestsPane();
  if(tab==='actions')renderActionsPane();
  if(tab==='marketing')renderMarketingPane();
  if(tab==='templates')renderTemplatesPane();
  if(tab==='reports')renderReportsPane();
  if(tab==='plan')renderPlanPane();
  if(tab==='import')renderImportPane();
  if(tab==='admin')renderAdminPane();
}



// ============================================================================
// PROPERTY SELECTOR — affects all data queries across tabs
// ============================================================================
// (duplicate renderPropertySelector removed)

async function getGuestIdsForProperty() {
  // Returns null = all guests, or array of guest IDs for specific property
  if (STATE.selectedPropertyId === null) return null;
  const propIds = [STATE.selectedPropertyId];
  let allStays = [], page = 0;
  while(true) {
    const {data:batch} = await sb.from('stays').select('guest_id').in('property_id', propIds).range(page*1000,(page+1)*1000-1);
    if(!batch||batch.length===0) break;
    allStays = allStays.concat(batch);
    if(batch.length < 1000) break;
    page++;
  }
  const data = allStays;
  return [...new Set((data || []).map(s => s.guest_id))];
}

// ============================================================================
// OVERVIEW
// ============================================================================
async function renderOverview(){
  const el=document.getElementById('pane-overview');
  el.innerHTML='<div class="loading">Loading</div>';

  const propIds = STATE.selectedPropertyId !== null
    ? [STATE.selectedPropertyId]
    : STATE.accessibleProperties.map(p=>p.id);
  if(!propIds.length){el.innerHTML='<div class="empty">No properties.</div>';return;}

  const [statsRes, inhouseRes, checkoutsRes, bdayGuestsRes] = await Promise.all([
    sb.rpc('get_overview_stats', {prop_ids: propIds}),
    sb.rpc('get_inhouse_guests', {prop_ids: propIds}),
    sb.rpc('get_recent_checkouts', {prop_ids: propIds, days_back: 3}),
    // Fetch guests with DOB for birthday calculation
    sb.from('stays').select('guests!inner(id,full_name,email,date_of_birth)')
      .in('property_id', propIds)
      .not('guests.date_of_birth', 'is', null)
      .not('guests.email', 'is', null)
      .then(r => r.error ? {data:[]} : r)
  ]);

  const stats = statsRes.data || {};
  const inhouse = inhouseRes.data || [];
  const checkouts = checkoutsRes.data || [];

  // Calculate days_until birthday in JS from raw guest data
  const _today = new Date(); _today.setHours(0,0,0,0);
  const _seen = new Set();
  const birthdays = (bdayGuestsRes.data||[])
    .map(s => s.guests || s)  // handle nested join or flat result
    .filter(g => g && g.id && g.date_of_birth && !_seen.has(g.id) && _seen.add(g.id))
    .map(g => {
      const dob = new Date(g.date_of_birth);
      let bday = new Date(_today.getFullYear(), dob.getMonth(), dob.getDate());
      if(bday < _today) bday = new Date(_today.getFullYear()+1, dob.getMonth(), dob.getDate());
      const days = Math.round((bday - _today) / 86400000);
      return {...g, days_until: days};
    })
    .filter(g => typeof g.days_until === 'number' && g.days_until >= 0 && g.days_until <= 30)
    .sort((a,b) => a.days_until - b.days_until);
  const scopeLabel = STATE.selectedPropertyId
    ? (STATE.properties.find(p=>p.id===STATE.selectedPropertyId)?.name||'Property')
    : 'All Properties';

  el.innerHTML=`
    <div class="sg">
      <div class="sm"><div class="sl">Total guests</div><div class="sv">${(stats.total_guests||0).toLocaleString()}</div><div class="ss">${escapeHtml(scopeLabel)}</div></div>
      <div class="sm"><div class="sl">In-house</div><div class="sv">${(stats.in_house||0).toLocaleString()}</div><div class="ss">Stayover + Checked IN</div></div>
      <div class="sm"><div class="sl">Repeat guests</div><div class="sv">${(stats.repeat_guests||0).toLocaleString()}</div></div>
      <div class="sm"><div class="sl">Direct bookings</div><div class="sv">${(stats.direct_bookings||0).toLocaleString()}</div></div>
      <div class="sm"><div class="sl">Total stays</div><div class="sv">${(stats.total_stays||0).toLocaleString()}</div></div>
      <div class="sm"><div class="sl">Net revenue</div><div class="sv">$${Math.round(stats.net_revenue||0).toLocaleString()}</div></div>
    </div>
    <div class="g2">
      <div class="card"><div class="ct" style="margin-bottom:12px">Today's urgent actions</div><div id="ov-alerts"></div></div>
      <div class="card"><div class="ct" style="margin-bottom:12px">Upcoming birthdays (30 days)</div><div id="ov-bdays"></div></div>
    </div>
    <div class="card"><div class="ct" style="margin-bottom:12px">Current in-house · ${escapeHtml(scopeLabel)}</div><div id="ov-inhouse"></div></div>`;

  // Alerts
  const items=[];
  checkouts.slice(0,3).forEach(g=>{items.push({dot:'#A32D2D',txt:fn(g.full_name)+' — send review request',sub:'Checked out '+fmtDate(g.departure_date)+' · '+(g.property_name||'')});});
  inhouse.filter(s=>s.nights>=3).slice(0,2).forEach(s=>items.push({dot:'var(--kaani-orange)',txt:fn(s.full_name)+' — upsell excursions',sub:'Until '+fmtDate(s.departure_date)+' · '+(s.property_name||'')}));
  document.getElementById('ov-alerts').innerHTML=items.length
    ?items.map(it=>`<div class="row"><div class="dot" style="background:${it.dot}"></div><div style="flex:1;min-width:0"><div class="rn">${escapeHtml(it.txt)}</div><div style="font-size:11px;color:var(--gray-text)">${escapeHtml(it.sub)}</div></div></div>`).join('')
    :'<div class="empty">No urgent actions.</div>';

  // Birthdays
  document.getElementById('ov-bdays').innerHTML=birthdays.length
    ?birthdays.slice(0,5).map(g=>`<div class="row"><div style="flex:1;min-width:0"><div class="rn">${escapeHtml(fn(g.full_name||''))}</div><div style="font-size:11px;color:var(--gray-text)">${escapeHtml(g.email||'')}</div></div><div class="rd" style="color:${g.days_until<=3?'var(--danger)':g.days_until<=7?'var(--warning)':'var(--success)'};font-weight:600">in ${g.days_until}d</div></div>`).join('')
    :'<div class="empty">No birthdays in next 30 days.</div>';

  // In-house list
  document.getElementById('ov-inhouse').innerHTML=inhouse.length
    ?inhouse.slice(0,10).map(s=>{const stCls=s.status?.toLowerCase().includes('stayover')?'bm':'be';return`<div class="row"><div class="av">${ini(s.full_name||'?')}</div><div style="flex:1;min-width:0"><div class="rn">${escapeHtml(s.full_name||'')}</div><div style="font-size:11px;color:var(--gray-text)">${escapeHtml(s.nationality||'')} · ${escapeHtml(s.room_number||'')} · until ${fmtDate(s.departure_date)} · ${escapeHtml(s.property_name||'')}</div></div><span class="badge ${stCls}">${escapeHtml(s.status||'')}</span></div>`;}).join('')
    +(inhouse.length>10?`<div style="font-size:12px;color:var(--gray-text);padding:10px 0">+${inhouse.length-10} more in-house</div>`:'')
    :'<div class="empty">No in-house guests.</div>';
}

async function renderDashboard(){
  const el=document.getElementById('pane-dashboard');
  el.innerHTML='<div class="loading">Loading dashboard</div>';

  const propIds = STATE.selectedPropertyId !== null
    ? [STATE.selectedPropertyId]
    : STATE.accessibleProperties.map(p=>p.id);
  if(!propIds.length){el.innerHTML='<div class="empty">No properties.</div>';return;}

  const [perfRes, statsRes, qualRes, srcRes, natRes, nightsRes, typeRes, arrRes] = await Promise.all([
    sb.from('v_property_performance').select('*'),
    sb.rpc('get_overview_stats', {prop_ids: propIds}),
    sb.rpc('get_data_quality', {prop_ids: propIds}),
    sb.rpc('get_source_breakdown', {prop_ids: propIds}),
    sb.rpc('get_nationality_breakdown', {prop_ids: propIds}),
    sb.rpc('get_nights_breakdown', {prop_ids: propIds}),
    sb.rpc('get_guest_type_breakdown', {prop_ids: propIds}),
    sb.rpc('get_daily_arrivals', {prop_ids: propIds})
  ]);

  const stats = statsRes.data || {};
  const q = qualRes.data || {};
  const propPerf = (perfRes.data || []).filter(p=>STATE.accessibleProperties.map(x=>x.id).includes(p.property_id));
  const scopeLabel = STATE.selectedPropertyId
    ? (STATE.properties.find(p=>p.id===STATE.selectedPropertyId)?.name||'Property')
    : 'All Properties';

  const pct = (n,d) => d>0 ? Math.round(1000*n/d)/10 : 0;

  const perfRows = propPerf.length>0 ? propPerf.map(p=>{
    const ep=p.unique_guests>0?Math.round(100*p.guests_with_email/p.unique_guests):0;
    const cur=STATE.selectedPropertyId===p.property_id;
    return`<tr style="border-bottom:1px solid var(--line-peach);${cur?'background:var(--kaani-cream-deep)':''}">
      <td style="padding:12px;font-weight:600"><div style="display:flex;align-items:center;gap:8px"><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:var(--kaani-orange)"></span>${escapeHtml(p.property_name)}</div></td>
      <td style="padding:12px;text-align:right;font-weight:600">${(p.unique_guests||0).toLocaleString()}</td>
      <td style="padding:12px;text-align:right">${(p.total_stays||0).toLocaleString()}</td>
      <td style="padding:12px;text-align:right">${p.avg_nights||'—'}</td>
      <td style="padding:12px;text-align:right">${p.in_house||0}</td>
      <td style="padding:12px;text-align:right">${p.direct_guests||0}</td>
      <td style="padding:12px;text-align:right;color:${ep>=80?'var(--success)':ep>=50?'var(--warning)':'var(--danger)'};font-weight:600">${ep}%</td>
    </tr>`;
  }).join('') : '<tr><td colspan="7" style="padding:20px;text-align:center;color:var(--gray-text)">Import guest data to see comparison</td></tr>';

  el.innerHTML=`
    <div class="card">
      <div class="ch"><div class="ct">Property Performance</div><div class="cs">all time comparison</div></div>
      <div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:13px;min-width:560px">
        <thead><tr style="border-bottom:1px solid var(--line-peach)">
          <th style="text-align:left;padding:10px 12px;font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--gray-text);font-weight:600">Property</th>
          <th style="text-align:right;padding:10px 12px;font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--gray-text);font-weight:600">Guests</th>
          <th style="text-align:right;padding:10px 12px;font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--gray-text);font-weight:600">Stays</th>
          <th style="text-align:right;padding:10px 12px;font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--gray-text);font-weight:600">Avg nights</th>
          <th style="text-align:right;padding:10px 12px;font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--gray-text);font-weight:600">In-house</th>
          <th style="text-align:right;padding:10px 12px;font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--gray-text);font-weight:600">Direct</th>
          <th style="text-align:right;padding:10px 12px;font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--gray-text);font-weight:600">Email %</th>
        </tr></thead>
        <tbody>${perfRows}</tbody>
      </table></div>
    </div>
    <div style="margin:20px 0 14px;padding-bottom:12px;border-bottom:1px solid var(--line-peach)">
      <div style="font-family:'Inter Tight',sans-serif;font-size:20px;font-weight:700;color:var(--black);letter-spacing:-0.02em">${escapeHtml(scopeLabel)} — Detailed View</div>
    </div>
    <div class="sg">
      <div class="kpi"><div class="sl">Unique guests</div><div class="sv">${(stats.total_guests||0).toLocaleString()}</div><div class="ss">${escapeHtml(scopeLabel)}</div><div class="kpi-bar"><div class="kpi-fill" style="width:100%;background:var(--kaani-orange)"></div></div></div>
      <div class="kpi"><div class="sl">Email capture</div><div class="sv">${pct(q.has_email,q.total_guests)}%</div><div class="ss">${(q.has_email||0).toLocaleString()} of ${(q.total_guests||0).toLocaleString()}</div><div class="kpi-bar"><div class="kpi-fill" style="width:${pct(q.has_email,q.total_guests)}%;background:#378ADD"></div></div></div>
      <div class="kpi"><div class="sl">Phone on file</div><div class="sv">${pct(q.has_phone,q.total_guests)}%</div><div class="ss">${(q.has_phone||0).toLocaleString()} of ${(q.total_guests||0).toLocaleString()}</div><div class="kpi-bar"><div class="kpi-fill" style="width:${pct(q.has_phone,q.total_guests)}%;background:#EF9F27"></div></div></div>
      <div class="kpi"><div class="sl">DOB on file</div><div class="sv">${pct(q.has_dob,q.total_guests)}%</div><div class="ss">${(q.has_dob||0).toLocaleString()} of ${(q.total_guests||0).toLocaleString()}</div><div class="kpi-bar"><div class="kpi-fill" style="width:${pct(q.has_dob,q.total_guests)}%;background:var(--success)"></div></div></div>
    </div>
    <div class="g2">
      <div class="card"><div class="ch"><div class="ct">Booking source</div><div class="cs">all stays</div></div><div class="cw"><canvas id="srcChart"></canvas></div></div>
      <div class="card"><div class="ch"><div class="ct">Top nationalities</div><div class="cs">guest count</div></div><div class="cw"><canvas id="natChart"></canvas></div></div>
    </div>
    <div class="g2">
      <div class="card"><div class="ch"><div class="ct">Stay length</div><div class="cs">stays by nights</div></div><div class="cw"><canvas id="nightsChart"></canvas></div></div>
      <div class="card"><div class="ch"><div class="ct">Tourist vs Local</div><div class="cs">guest type</div></div><div class="cw"><canvas id="typeChart"></canvas></div></div>
    </div>
    <div class="card"><div class="ch"><div class="ct">Daily arrivals</div><div class="cs">last 30 days</div></div><div class="cw" style="height:200px"><canvas id="arrChart"></canvas></div></div>
    <div class="card">
      <div class="ct" style="margin-bottom:12px">Data quality · ${escapeHtml(scopeLabel)}</div>
      <div style="font-size:13px;color:var(--gray-text)">
        <div style="display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid var(--line-peach)"><span>Email capture</span><strong style="color:var(--black)">${pct(q.has_email,q.total_guests)}%</strong></div>
        <div style="display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid var(--line-peach)"><span>Phone capture</span><strong style="color:var(--black)">${pct(q.has_phone,q.total_guests)}%</strong></div>
        <div style="display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid var(--line-peach)"><span>DOB capture</span><strong style="color:var(--black)">${pct(q.has_dob,q.total_guests)}%</strong></div>
        <div style="display:flex;justify-content:space-between;padding:8px 0"><span>ID / Passport</span><strong style="color:var(--black)">${pct(q.has_id,q.total_guests)}%</strong></div>
      </div>
    </div>
    <div class="card"><div class="ct" style="margin-bottom:12px">Key insights · ${escapeHtml(scopeLabel)}</div><div id="insights-body"><div class="loading">Generating</div></div></div>`;

  // Render all charts from RPC data
  const srcData = srcRes.data || [];
  const natData = natRes.data || [];
  const nightsData = nightsRes.data || [];
  const typeData = typeRes.data || [];
  const arrData = arrRes.data || [];

  new Chart(document.getElementById('srcChart'),{type:'doughnut',data:{labels:srcData.map(r=>srcShort(r.source)),datasets:[{data:srcData.map(r=>r.stay_count),backgroundColor:['#F47923','#2C7AB5','#1D9E75','#BA7517','#A32D2D','#6E5BB8','#EF9F27','#9FE1CB'],borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,cutout:'65%',plugins:{legend:{position:'bottom',labels:{font:{size:11},padding:12}}}}});

  const natColors=['#F47923','#2C7AB5','#1D9E75','#BA7517','#A32D2D','#6E5BB8','#EF9F27','#5DCAA5','#E8637A','#4A90D9','#8BC34A','#FF7043','#26C6DA','#AB47BC','#EC407A'];
  const natChartHeight = Math.max(240, natData.length * 28);
  document.getElementById('natChart').parentElement.style.height = natChartHeight+'px';
  new Chart(document.getElementById('natChart'),{
    type:'bar',
    data:{
      labels:natData.map(r=>r.nationality),
      datasets:[{
        data:natData.map(r=>r.guest_count),
        backgroundColor:natData.map((_,i)=>natColors[i%natColors.length]),
        borderRadius:4
      }]
    },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      indexAxis:'y',
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>' '+ctx.raw.toLocaleString()+' guests'}}},
      scales:{
        x:{beginAtZero:true,ticks:{font:{size:11}}},
        y:{ticks:{font:{size:11},autoSkip:false}}
      }
    }
  });

  new Chart(document.getElementById('nightsChart'),{type:'bar',data:{labels:nightsData.map(r=>r.nights+'n'),datasets:[{data:nightsData.map(r=>r.stay_count),backgroundColor:'#FBD7BC',borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{font:{size:11}}},x:{ticks:{font:{size:11}}}}}});

  const tourist = typeData.find(r=>r.guest_type?.toLowerCase().includes('tourist'))?.guest_count||0;
  const local = typeData.find(r=>r.guest_type?.toLowerCase().includes('local'))?.guest_count||0;
  const totalT = tourist+local;
  new Chart(document.getElementById('typeChart'),{type:'doughnut',data:{labels:['Tourist','Local'],datasets:[{data:[tourist,local],backgroundColor:['#F47923','#FBD7BC'],borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,cutout:'65%',plugins:{legend:{position:'bottom',labels:{font:{size:11},padding:12,generateLabels:(c)=>{const d=c.data;return d.labels.map((l,i)=>({text:l+' '+d.datasets[0].data[i]+(totalT?' ('+Math.round(d.datasets[0].data[i]/totalT*100)+'%)':''),fillStyle:d.datasets[0].backgroundColor[i],strokeStyle:'transparent',index:i}));}}}}}});

  new Chart(document.getElementById('arrChart'),{type:'line',data:{labels:arrData.map(r=>{const dt=new Date(r.arrival_date);return String(dt.getDate()).padStart(2,'0')+' '+MONTHS[dt.getMonth()];}),datasets:[{data:arrData.map(r=>r.arrival_count),borderColor:'#F47923',backgroundColor:'rgba(244,121,35,0.1)',fill:true,tension:0.3,pointBackgroundColor:'#F47923',pointRadius:4,pointHoverRadius:6,borderWidth:2.5}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{font:{size:11}}},x:{ticks:{font:{size:10},maxRotation:0}}}}});

  // Insights
  const sourceCounts = Object.fromEntries(srcData.map(r=>[r.source,+r.stay_count]));
  const natCounts = Object.fromEntries(natData.map(r=>[r.nationality,+r.guest_count]));
  const nightsDist = Object.fromEntries(nightsData.map(r=>[r.nights,+r.stay_count]));
  const typeCounts = {tourist, local};
  generateInsights(stats, {total_guests:q.total_guests,pct_email:pct(q.has_email,q.total_guests)}, sourceCounts, natCounts, typeCounts, nightsDist);
}



function generateInsights(stats, quality, sources, nats, types, nightsDist){
  const el = document.getElementById('insights-body');
  if(!el) return;
  const insights = [];
  const totalStays = Object.values(sources).reduce((a,b)=>a+b, 0);
  const totalGuests = Object.values(nats).reduce((a,b)=>a+b, 0);

  // Source concentration
  if(totalStays > 0){
    const sorted = Object.entries(sources).sort((a,b)=>b[1]-a[1]);
    const top = sorted[0];
    const pct = Math.round(top[1]/totalStays*100);
    if(pct > 35) insights.push({
      title:'Source concentration risk.',
      body:`${pct}% of stays come through ${srcShort(top[0])}. Diversifying reduces OTA dependency and saves commission costs.`
    });
  }

  // Nationality dominance
  if(totalGuests > 0){
    const sorted = Object.entries(nats).sort((a,b)=>b[1]-a[1]);
    const top = sorted[0];
    const pct = Math.round(top[0][1]/totalGuests*100);
    if(sorted.length > 0){
      const topNat = sorted[0][0];
      const topPct = Math.round(sorted[0][1]/totalGuests*100);
      if(topPct > 25) insights.push({
        title:`${topNat} market dominance.`,
        body:`${topPct}% of guests are from ${topNat}. Consider targeted campaigns and culturally tailored welcome touches.`
      });
    }
  }

  // Long stay opportunity
  const totalN = Object.values(nightsDist).reduce((a,b)=>a+b, 0);
  const longStays = Object.entries(nightsDist).filter(([n])=>+n>=7).reduce((a,[,v])=>a+v, 0);
  if(totalN > 0 && longStays/totalN > 0.25) insights.push({
    title:'Long-stay loyalty opportunity.',
    body:`${Math.round(longStays/totalN*100)}% of stays are 7+ nights — these guests are your highest-value win-back targets.`
  });

  // Email gap
  const emailPct = quality.pct_email || 0;
  if(emailPct < 90) insights.push({
    title:'Email capture gap.',
    body:`${emailPct}% of guests have an email on file. Every extra 1% here means more reach for all campaigns.`
  });

  // Direct booking gap
  const directCount = Object.entries(sources).filter(([s])=>s&&s.toUpperCase().includes('DIRECT')).reduce((a,[,v])=>a+v, 0);
  if(totalStays > 0){
    const directPct = Math.round(directCount/totalStays*100);
    if(directPct < 20) insights.push({
      title:'Direct booking upside.',
      body:`Only ${directPct}% of bookings are direct. Industry benchmark is 20-25%. The win-back campaign in your Plan tab targets this directly.`
    });
  }

  el.innerHTML = insights.length
    ? insights.map(i=>`<div class="insight-item"><strong>${i.title}</strong> ${i.body}</div>`).join('')
    : '<div style="font-size:13px;color:var(--gray-text);padding:10px 0">Import data across more properties to generate meaningful insights.</div>';
}



function filterNoEmail(){
  // Show guests with no email — use fallback direct query
  document.getElementById('glist').innerHTML='<div class="loading">Loading guests without email</div>';
  document.getElementById('g-pagination').innerHTML='';
  const propIds = STATE.selectedPropertyId !== null
    ? [STATE.selectedPropertyId]
    : STATE.accessibleProperties.map(p=>p.id);
  
  // Get all guest IDs for this property
  sb.from('stays').select('guest_id').in('property_id', propIds)
    .then(async r => {
      const gIds = [...new Set((r.data||[]).map(s=>s.guest_id))];
      // Fetch guests with no email in chunks
      let all = [];
      for(let i=0;i<gIds.length;i+=500){
        const chunk = gIds.slice(i,i+500);
        const{data} = await sb.from('guests').select('*').in('id',chunk).is('email',null);
        if(data) all = all.concat(data);
      }
      STATE.currentGuests = all;
      const summary = document.getElementById('g-filter-summary');
      if(summary){
        summary.innerHTML=`<div style="font-size:13px;color:var(--kaani-orange);font-weight:600;padding:8px 0">${all.length.toLocaleString()} guests have no email — click a guest to add their email manually</div>`;
        summary.style.display='block';
      }
      renderGuestList(all.length);
    });
}


async function renderGuestsPane(){
  const el=document.getElementById('pane-guests');
  el.innerHTML=`
    <div class="sg" id="g-stats"></div>
    <div class="toolbar">
      <input type="text" id="gsearch" placeholder="Search name, email, nationality..." />
      <select id="gfst"><option value="">All statuses</option><option value="prospect">Prospect</option><option value="first_time">First-time</option><option value="repeat">Repeat</option><option value="vip">VIP</option><option value="lapsed">Lapsed</option></select>
      <select id="gfsrc"><option value="">All sources</option></select>
      <select id="gfnat"><option value="">All nationalities</option></select>
      <select id="gfbday"><option value="">All guests</option><option value="7">Birthday in 7 days</option><option value="30">Birthday in 30 days</option><option value="60">Birthday in 60 days</option></select>
    </div>
    <div id="g-filter-summary" style="display:none"></div>
    <div id="glist"></div>
    <div class="pagination" id="g-pagination"></div>`;

  let qual;
  const gIds = await getGuestIdsForProperty();
  if(gIds === null){
    const{data}=await sb.from('v_data_quality').select('*').single();
    qual = data;
  } else if(gIds.length === 0){
    qual = {total_guests:0,has_email:0,pct_email:0,has_phone:0,pct_phone:0,has_dob:0,pct_dob:0};
  } else {
    const g = (await sb.from('guests').select('email,phone,date_of_birth').in('id', gIds)).data || [];
    const total = g.length || 1;
    qual = {total_guests:g.length,has_email:g.filter(x=>x.email).length,pct_email:Math.round(g.filter(x=>x.email).length/total*1000)/10,has_phone:g.filter(x=>x.phone).length,pct_phone:Math.round(g.filter(x=>x.phone).length/total*1000)/10,has_dob:g.filter(x=>x.date_of_birth).length,pct_dob:Math.round(g.filter(x=>x.date_of_birth).length/total*1000)/10};
  }
  document.getElementById('g-stats').innerHTML=`
    <div class="sm"><div class="sl">Total</div><div class="sv">${(qual?.total_guests||0).toLocaleString()}</div></div>
    <div class="sm"><div class="sl">With email</div><div class="sv">${qual?.has_email||0}</div><div class="ss">${qual?.pct_email||0}%</div></div>
    <div class="sm"><div class="sl">With phone</div><div class="sv">${qual?.has_phone||0}</div><div class="ss">${qual?.pct_phone||0}%</div></div>
    <div class="sm"><div class="sl">With DOB</div><div class="sv">${qual?.has_dob||0}</div><div class="ss">${qual?.pct_dob||0}%</div></div>`;

  // Populate nationality + source filters
  // Use RPC nationality breakdown — already has all nationalities, no row limit
  const[natsRpc, sourcesRes]=await Promise.all([
    sb.rpc('get_nationality_breakdown', {prop_ids: [1,2,3,4,5]}),
    sb.from('stays').select('source').not('source','is',null)
  ]);
  // Build natsRes compatible structure from RPC
  const natsRes = {data: (natsRpc.data||[]).map(r=>({nationality:r.nationality}))};
  const uniq=[...new Set((natsRes.data||[]).map(n=>n.nationality))].sort();
  const ne=document.getElementById('gfnat');
  if(ne){uniq.forEach(n=>{const o=document.createElement('option');o.value=n;o.text=n;ne.appendChild(o);});}
  const uniqSrc=[...new Set((sourcesRes.data||[]).map(s=>s.source))].sort();
  const se=document.getElementById('gfsrc');
  if(se){uniqSrc.forEach(s=>{const o=document.createElement('option');o.value=s;o.text=srcShort(s);se.appendChild(o);});}

  document.getElementById('gsearch').oninput=()=>{STATE.guestsPage=0;loadGuests();};
  document.getElementById('gfst').onchange=()=>{STATE.guestsPage=0;loadGuests();};
  document.getElementById('gfsrc').onchange=()=>{STATE.guestsPage=0;loadGuests();};
  document.getElementById('gfnat').onchange=()=>{STATE.guestsPage=0;loadGuests();};
  document.getElementById('gfbday').onchange=()=>{STATE.guestsPage=0;loadGuests();};

  loadGuests();
}

let guestSearchTimer;
async function loadGuests(){
  clearTimeout(guestSearchTimer);
  guestSearchTimer=setTimeout(async()=>{
    const q     = document.getElementById('gsearch')?.value.trim()||'';
    const fst   = document.getElementById('gfst')?.value||'';
    const fnat  = document.getElementById('gfnat')?.value||'';
    const fsrc  = document.getElementById('gfsrc')?.value||'';
    const fbday = document.getElementById('gfbday')?.value||'';

    const propIds = STATE.selectedPropertyId !== null
      ? [STATE.selectedPropertyId]
      : STATE.accessibleProperties.map(p=>p.id);

    const from = STATE.guestsPage * STATE.guestsPerPage;

    // Use server-side RPC for property-filtered search (no URL length issues)
    const{data, error} = await sb.rpc('search_guests', {
      prop_ids: propIds,
      search_text: q,
      filter_status: fst,
      filter_nat: fnat,
      filter_source: fsrc,
      page_from: from,
      page_size: STATE.guestsPerPage
    });

    if(error){
      console.warn('search_guests RPC error:', error.message, 'Falling back...');
      await loadGuestsFallback(q, fst, fnat, fsrc, fbday, from);
      return;
    }

    let guests = data || [];
    const totalCount = guests.length > 0 ? Number(guests[0].total_count) : 0;
    console.log('search_guests returned:', guests.length, 'guests, total:', totalCount, 'propIds:', propIds);

    // Birthday filter — client side
    if(fbday && +fbday > 0){
      const today = new Date(); today.setHours(0,0,0,0);
      guests = guests.filter(g => {
        if(!g.date_of_birth) return false;
        const dob = new Date(g.date_of_birth);
        let bday = new Date(today.getFullYear(), dob.getMonth(), dob.getDate());
        if(bday < today) bday = new Date(today.getFullYear()+1, dob.getMonth(), dob.getDate());
        const days = Math.round((bday-today)/86400000);
        return days >= 0 && days <= +fbday;
      });
    }

    STATE.currentGuests = guests;
    renderGuestList(fbday ? guests.length : totalCount);
  }, 200);
}

async function loadGuestsFallback(q, fst, fnat, fsrc, fbday, from){
  // Fallback for when search_guests RPC is not installed
  const to = from + STATE.guestsPerPage - 1;
  let query = sb.from('guests').select('*',{count:'exact'})
    .order('last_stay_date',{ascending:false,nullsFirst:false})
    .range(from, to);
  if(q) query = query.or(`full_name.ilike.%${q}%,email.ilike.%${q}%,nationality.ilike.%${q}%`);
  if(fst) query = query.eq('lead_status', fst);
  if(fnat) query = query.eq('nationality', fnat);
  const{data, error, count} = await query;
  if(error){toast('Load error: '+error.message, true);return;}
  let filtered = data||[];
  if(fbday && +fbday>0){
    const today=new Date();today.setHours(0,0,0,0);
    filtered=filtered.filter(g=>{
      if(!g.date_of_birth)return false;
      const dob=new Date(g.date_of_birth);
      let bday=new Date(today.getFullYear(),dob.getMonth(),dob.getDate());
      if(bday<today)bday=new Date(today.getFullYear()+1,dob.getMonth(),dob.getDate());
      return Math.round((bday-today)/86400000)<=+fbday;
    });
  }
  STATE.currentGuests = filtered;
  renderGuestList(fbday?filtered.length:(count||0));
}



function renderGuestList(totalCount){
  const el=document.getElementById('glist');
  // Show active filter summary
  const filterSummary = document.getElementById('g-filter-summary');
  if(filterSummary){
    const fst=document.getElementById('gfst')?.value;
    const fnat=document.getElementById('gfnat')?.value;
    const fsrc=document.getElementById('gfsrc')?.value;
    const fbday=document.getElementById('gfbday')?.value;
    const q=document.getElementById('gsearch')?.value?.trim();
    const active = [fst,fnat,fsrc,fbday,q].filter(Boolean);
    if(active.length > 0){
      filterSummary.innerHTML=`<div style="font-size:13px;color:var(--kaani-orange);font-weight:600;padding:8px 0">${totalCount.toLocaleString()} guests match your filters</div>`;
      filterSummary.style.display='block';
    } else {
      filterSummary.style.display='none';
    }
  }
  if(!STATE.currentGuests.length){el.innerHTML='<div class="empty">No guests match your filters</div>';document.getElementById('g-pagination').innerHTML='';return;}
  el.innerHTML=STATE.currentGuests.map(g=>{
    const age=calcAge(g.date_of_birth);
    const today=new Date();today.setHours(0,0,0,0);
    let bdayBadge='';
    if(g.date_of_birth){
      const dob=new Date(g.date_of_birth);
      let bday=new Date(today.getFullYear(),dob.getMonth(),dob.getDate());
      if(bday<today)bday=new Date(today.getFullYear()+1,dob.getMonth(),dob.getDate());
      const days=Math.round((bday-today)/86400000);
      if(days<=30)bdayBadge=` <span style="color:var(--kaani-orange);font-weight:600">🎂 in ${days}d</span>`;
    }
    const stCls=g.lead_status==='vip'?'bu':g.lead_status==='repeat'?'be':g.lead_status==='lapsed'?'bg':'bm';
    return`<div class="gcard" onclick="openGuestPanel('${g.id}')">
      <div style="display:flex;align-items:center;gap:11px">
        <div class="av">${ini(g.full_name)}</div>
        <div style="flex:1;min-width:0">
          <div style="font-size:14px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${escapeHtml(g.full_name)}</div>
          <div style="font-size:12px;color:var(--gray-text)">${g.date_of_birth?'DOB: '+fmtDate(g.date_of_birth)+(age?' · age '+age:'')+''+bdayBadge:escapeHtml(g.email||'no email')}</div>
        </div>
        <div style="display:flex;flex-direction:column;align-items:flex-end;gap:4px;flex-shrink:0">
          <span class="badge ${stCls}">${(g.lead_status||'').replace('_',' ')}</span>
          <span style="font-size:11px;color:var(--gray-text)">${g.total_stays} ${g.total_stays===1?'stay':'stays'}</span>
        </div>
      </div>
      <div style="display:flex;gap:10px;margin-top:9px;padding-top:9px;border-top:1px solid var(--line-peach);flex-wrap:wrap;font-size:12px">
        <span style="color:var(--gray-text)"><strong style="color:var(--black)">${escapeHtml(g.nationality||'—')}</strong></span>
        <span style="color:var(--gray-text)">Last stay <strong style="color:var(--black)">${g.last_stay_date?fmtDate(g.last_stay_date):'—'}</strong></span>
        ${g.guest_type==='local'?'<span class="badge bl">Local</span>':''}
      </div>
    </div>`;
  }).join('');
  renderPagination(totalCount);
}

function renderPagination(total){
  const pgEl=document.getElementById('g-pagination');
  if(!pgEl)return;
  const totalPages=Math.ceil(total/STATE.guestsPerPage);
  if(totalPages<=1){pgEl.innerHTML='';return;}
  const p=STATE.guestsPage;
  let html=`<button onclick="changeGuestPage(${p-1})" ${p===0?'disabled':''}>‹ Prev</button>`;
  const start=Math.max(0,p-2);const end=Math.min(totalPages-1,p+2);
  for(let i=start;i<=end;i++)html+=`<button class="${i===p?'on':''}" onclick="changeGuestPage(${i})">${i+1}</button>`;
  html+=`<button onclick="changeGuestPage(${p+1})" ${p>=totalPages-1?'disabled':''}>Next ›</button>`;
  html+=`<span style="font-size:11px;color:var(--gray-text);align-self:center;margin-left:8px">Page ${p+1} of ${totalPages} · ${total.toLocaleString()} guests</span>`;
  pgEl.innerHTML=html;
}

function changeGuestPage(p){STATE.guestsPage=p;loadGuests();}



async function editGuestEmail(guestId, currentEmail){
  const newEmail = prompt('Enter email address:', currentEmail||'');
  if(newEmail === null) return; // cancelled
  const cleaned = newEmail.trim().toLowerCase();
  if(cleaned && !cleaned.includes('@')){toast('Please enter a valid email address',true);return;}
  const{error}=await sb.from('guests').update({email:cleaned||null}).eq('id',guestId);
  if(error){toast(error.message,true);return;}
  toast(cleaned?'Email updated':'Email removed');
  openGuestPanel(guestId);
}


async function openGuestPanel(id){
  const{data:g}=await sb.from('guests').select('*').eq('id',id).single();
  if(!g)return;
  const{data:stays}=await sb.from('stays').select('*,properties(name)').eq('guest_id',id).order('arrival_date',{ascending:false});
  const{data:gtags}=await sb.from('guest_tags').select('*,tags(*)').eq('guest_id',id);
  const{data:comms}=await sb.from('communications').select('*').eq('guest_id',id).order('sent_at',{ascending:false}).limit(5);

  const[bg,fg]=avc(g.full_name);const age=calcAge(g.date_of_birth);
  const sp=document.getElementById('sp');
  sp.style.display='block';
  sp.innerHTML=`<div class="panel-wrap" onclick="closeGuestPanel(event)">
    <div class="panel" onclick="event.stopPropagation()">
      <button class="pc" onclick="closeGuestPanelDirect()">✕</button>
      <div class="pav" style="background:${bg};color:${fg}">${ini(g.full_name)}</div>
      <div class="pn">${escapeHtml(g.full_name)}</div>
      <div class="ps">${escapeHtml(g.nationality||'')}${age?' · age '+age:''} · ${g.guest_type==='local'?'Local':'Tourist'} · ${g.lead_status.replace('_',' ')}</div>

      <div style="display:flex;gap:8px;margin-bottom:16px">
        ${g.email?`<button class="btn btnp" style="flex:1" onclick="copyToClipboard('${escapeHtml(g.email)}','Email copied')">Copy email</button>`:'<button class="btn" style="flex:1" disabled>No email</button>'}
        <button class="btn" onclick="addTagPrompt('${id}')">+ Tag</button>
      </div>

      <div style="margin-bottom:16px">
        <div class="sec">Lifecycle</div>
        <div class="dr"><span class="dk">Total stays</span><span class="dv">${g.total_stays}</span></div>
        <div class="dr"><span class="dk">Total nights</span><span class="dv">${g.total_nights}</span></div>
        <div class="dr"><span class="dk">Total revenue</span><span class="dv">$${(g.total_revenue_usd||0).toFixed(2)}</span></div>
        <div class="dr"><span class="dk">First stay</span><span class="dv">${g.first_stay_date?fmtDate(g.first_stay_date):'—'}</span></div>
        <div class="dr"><span class="dk">Last stay</span><span class="dv">${g.last_stay_date?fmtDate(g.last_stay_date):'—'}</span></div>
      </div>

      <div style="margin-bottom:16px">
        <div class="sec">Profile</div>
        ${g.date_of_birth?`<div class="dr"><span class="dk">Date of birth</span><span class="dv dob">${fmtDate(g.date_of_birth)}</span></div>`:''}
        ${age?`<div class="dr"><span class="dk">Age</span><span class="dv">${age} years</span></div>`:''}
        <div class="dr"><span class="dk">Nationality</span><span class="dv">${escapeHtml(g.nationality||'—')}</span></div>
        <div class="dr"><span class="dk">Email</span>
          <span class="dv" style="display:flex;align-items:center;gap:8px">
            ${g.email
              ? escapeHtml(g.email)
              : `<span style="color:var(--kaani-orange);font-style:italic">No email</span>`}
            <button style="font-size:10px;padding:3px 8px;border-radius:6px;border:1px solid var(--line-peach);background:var(--kaani-cream-deep);cursor:pointer;flex-shrink:0" onclick="editGuestEmail('${g.id}','${escapeHtml(g.email||'')}')">✏️ Edit</button>
          </span>
        </div>
        <div class="dr"><span class="dk">Phone</span><span class="dv">${escapeHtml(g.phone||'—')}</span></div>
        <div class="dr"><span class="dk">Passport</span><span class="dv">${escapeHtml(g.passport_number||'—')}</span></div>
        <div class="dr"><span class="dk">National ID</span><span class="dv">${escapeHtml(g.national_id||'—')}</span></div>
        <div class="dr"><span class="dk">Marketing OK</span><span class="dv">${g.marketing_consent?'Yes':'No (unsubscribed)'}</span></div>
      </div>

      <div style="margin-bottom:16px">
        <div class="sec">Stay history (${stays?.length||0})</div>
        ${(stays||[]).map(s=>`<div class="dr"><span class="dk">${fmtDate(s.arrival_date)} → ${fmtDate(s.departure_date)}</span><span class="dv">${escapeHtml(s.properties?.name||'')} · ${s.nights}n</span></div>`).join('')||'<div style="font-size:12px;color:#5F5E5A">No stays yet</div>'}
      </div>

      <div style="margin-bottom:16px">
        <div class="sec">Tags</div>
        <div>${(gtags||[]).length?(gtags||[]).map(t=>`<span class="tag-pill" style="background:${t.tags.color}20;color:${t.tags.color}">${escapeHtml(t.tags.name)} <span class="x" onclick="removeTag('${id}',${t.tag_id})">✕</span></span>`).join(''):'<span style="font-size:12px;color:#5F5E5A">No tags yet</span>'}</div>
      </div>

      <div style="margin-bottom:16px">
        <div class="sec">Recent communications (${comms?.length||0})</div>
        ${(comms||[]).map(c=>`<div class="dr"><span class="dk">${fmtDate(c.sent_at)}</span><span class="dv">${escapeHtml(c.channel)} · ${escapeHtml(c.subject||c.body?.slice(0,40)||'')}</span></div>`).join('')||'<div style="font-size:12px;color:#5F5E5A">No communications logged</div>'}
      </div>

      <div>
        <div class="sec">Add note</div>
        <textarea id="ni" style="width:100%;min-height:60px;font-size:13px;border:1px solid #E8E6E0;border-radius:8px;padding:9px;font-family:inherit;background:#fff;resize:vertical" placeholder="Note about this guest..."></textarea>
        <button class="btn" style="margin-top:6px;width:100%" onclick="saveGuestNote('${id}')">Log note</button>
      </div>
    </div>
  </div>`;
}

function closeGuestPanel(e){if(e&&!e.target.classList.contains('panel-wrap'))return;closeGuestPanelDirect();}
function closeGuestPanelDirect(){document.getElementById('sp').style.display='none';loadGuests();}

async function copyToClipboard(text,msg){
  await navigator.clipboard.writeText(text);
  toast(msg||'Copied to clipboard');
}

async function addTagPrompt(guestId){
  const tagNames=STATE.tags.map(t=>t.name).join(', ');
  const choice=prompt(`Tag name (${tagNames}) or new:`);
  if(!choice||!choice.trim())return;
  const trimmed=choice.trim();
  let tag=STATE.tags.find(t=>t.name.toLowerCase()===trimmed.toLowerCase());
  if(!tag){
    const{data,error}=await sb.from('tags').insert({name:trimmed}).select().single();
    if(error){toast('Could not create tag (admin only?): '+error.message,true);return;}
    tag=data;STATE.tags.push(tag);
  }
  const{error}=await sb.from('guest_tags').insert({guest_id:guestId,tag_id:tag.id,added_by:STATE.user.id});
  if(error&&!error.message.includes('duplicate')){toast(error.message,true);return;}
  toast('Tag added');
  openGuestPanel(guestId);
}

async function removeTag(guestId,tagId){
  await sb.from('guest_tags').delete().eq('guest_id',guestId).eq('tag_id',tagId);
  openGuestPanel(guestId);
}

async function saveGuestNote(guestId){
  const note=document.getElementById('ni').value.trim();
  if(!note)return;
  const{error}=await sb.from('communications').insert({guest_id:guestId,channel:'note',direction:'internal',body:note,staff_user_id:STATE.user.id,staff_name:STATE.profile.full_name});
  if(error){toast(error.message,true);return;}
  toast('Note saved');
  openGuestPanel(guestId);
}

// ============================================================================
// ACTIONS — uses email_templates from database
// ============================================================================
const WA_TEMPLATES = {
  review_request: `Hi [Name]! 👋 Thank you for staying at {{property_name}}. We hope you had a wonderful time in Maafushi! 🌊\n\nCould you spare 2 minutes to leave us a Google review? It means the world to our small team 🙏\n\n[YOUR GOOGLE REVIEW LINK]\n\nThank you! 😊\nKaani Hotels`,
  winback: `Hi [Name]! 👋 We loved having you at {{property_name}} — hope Maafushi left you with beautiful memories! 🌴\n\nSpecial offer just for you — book direct and save 10%:\n\n📎 [DIRECT BOOKING LINK]\n⏰ Valid until: [DATE]\n\nReply to check availability 😊\nKaani Hotels`,
  upsell: `Hi [Name]! 🌺 Hope you're loving your stay!\n\nWe can arrange:\n🐠 Island hopping & snorkelling — $30/pp\n💆 Relaxation massage — $50/pp\n🍽️ Lunch + dinner upgrade — $25/pp/day\n🐬 Sunset dolphin cruise — $35/pp\n\nJust reply here or pop by reception! 😊\nKaani Hotels`,
  birthday: `Hi [Name]! 🎂🎉 Wishing you a very happy birthday from all of us at {{property_name}}! 🌺\n\nAs a birthday gift — complimentary room upgrade on your next direct booking. Just mention this message when you book 🎁\n\nHope your day is as beautiful as the Maldives! 🌊\nKaani Hotels`,
  seasonal: `Hi [Name]! 🌴 Hope you have wonderful memories of your time with us in Maafushi!\n\nExclusive offer for past guests:\n✨ [OFFER DESCRIPTION]\n📅 Travel: [DATES]\n⏰ Book by: [DEADLINE]\n📎 [DIRECT BOOKING LINK]\n\nReply to check availability 😊\nKaani Hotels`,
  locals: `Assalamu Alaikum [Name]! 🌺\n\nKaani Hotels geon special fareeh ingey — local guestunnah special rates!\n\n✨ [SPECIAL FARE DETAILS]\n📅 Travel: [DATES]\n\nDhivehi rayyithunnah special rate eh arrange kohdhevey. Book kohlan:\n📎 [DIRECT BOOKING LINK]\n\nReply kurey nuvatha call kurey: +960 [NUMBER]\nKaani Hotels Team`
};


const CAMPAIGN_CODES = {
  review_request: 'RR',
  winback: 'WB',
  upsell: 'UP',
  birthday: 'BD',
  seasonal: 'SB',
  locals: 'LC'
};

function generateCampaignCode(key, propId){
  const prop = STATE.properties.find(p=>p.id===propId||p.id===parseInt(propId));
  const propCode = prop?.code || 'KNI';
  const typeCode = CAMPAIGN_CODES[key] || 'CM';
  const now = new Date();
  const dateCode = String(now.getDate()).padStart(2,'0') + String(now.getMonth()+1).padStart(2,'0');
  return `${propCode}-${typeCode}-${dateCode}`;
}


async function renderActionsPane(){
  const el=document.getElementById('pane-actions');
  const propLabel = STATE.selectedPropertyId
    ? (STATE.properties.find(p=>p.id===STATE.selectedPropertyId)?.name||'Property')
    : 'All Properties';
  el.innerHTML=`
    <div style="margin-bottom:14px;font-size:12px;color:var(--gray-text);font-weight:600;text-transform:uppercase;letter-spacing:.06em">
      Viewing: <span style="color:var(--kaani-orange)">${escapeHtml(propLabel)}</span>
    </div>
    <div class="nav" style="margin-bottom:14px" id="action-nav">
      <button class="nb on" onclick="renderAction('review_request',this)">Review request</button>
      <button class="nb" onclick="renderAction('winback',this)">Win-back</button>
      <button class="nb" onclick="renderAction('upsell',this)">Upsell</button>
      <button class="nb" onclick="renderAction('birthday',this)">Birthdays</button>
      <button class="nb" onclick="renderAction('seasonal',this)">Seasonal</button>
      <button class="nb" onclick="renderChecklist(this)">Checklist</button>
      <button class="nb" onclick="renderAction('locals',this)">🇲🇻 Locals</button>
    </div>
    <div id="action-content"><div class="loading">Loading</div></div>`;
  renderAction('review_request');
}

function renderChecklist(btn){
  if(btn){document.querySelectorAll('#action-nav .nb').forEach(b=>b.classList.remove('on'));btn.classList.add('on');}
  const el=document.getElementById('action-content');
  let html='<div class="card">';
  ['daily','weekly','monthly'].forEach((freq,fi)=>{
    const label=freq==='daily'?'Daily (5 min)':freq==='weekly'?'Weekly (15 min)':'Monthly (30 min)';
    html+='<div style="font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#5F5E5A;font-weight:500;margin:'+(fi===0?'0':'14px')+' 0 10px">'+label+'</div>';
    CHECKLIST_DATA[freq].forEach((it,i)=>{
      html+='<div style="display:flex;align-items:flex-start;gap:8px;margin-bottom:8px;font-size:13px"><input type="checkbox" id="ck-'+freq+'-'+i+'" style="margin-top:3px;accent-color:#1D9E75;width:15px;height:15px;cursor:pointer"><label for="ck-'+freq+'-'+i+'" style="cursor:pointer;line-height:1.5;flex:1">'+it+'</label></div>';
    });
    if(fi<2)html+='<div style="height:1px;background:#E8E6E0;margin:12px 0"></div>';
  });
  html+='</div>';
  el.innerHTML=html;
}

async function renderAction(key,btn){
  if(btn){document.querySelectorAll('#action-nav .nb').forEach(b=>b.classList.remove('on'));btn.classList.add('on');}
  const tpl=STATE.templates.find(t=>t.template_key===key);
  if(!tpl){document.getElementById('action-content').innerHTML='<div class="empty">Template not found</div>';return;}

  const propIds = STATE.selectedPropertyId!==null
    ? [STATE.selectedPropertyId]
    : STATE.accessibleProperties.map(p=>p.id);
  const tplPropName = STATE.selectedPropertyId
    ? (STATE.properties.find(p=>p.id===STATE.selectedPropertyId)?.name||'Kaani Hotels')
    : 'Kaani Hotels';

  let guests=[];
  let segmentLabel='';

  if(key==='review_request'){
    const{data}=await sb.rpc('get_recent_checkouts',{prop_ids:propIds,days_back:3});
    guests=(data||[]);
    segmentLabel=`Recent checkouts (last 3 days) · ${tplPropName}`;
  }else if(key==='winback'){
    const gIds=await getGuestIdsForProperty();
    let q=sb.from('guests').select('id,full_name,email,marketing_consent,lead_status').not('email','is',null).eq('marketing_consent',true).in('lead_status',['first_time','lapsed']).order('last_stay_date',{ascending:false});
    if(gIds){if(gIds.length===0){guests=[];}else{q=q.in('id',gIds);}}
    if(!gIds||gIds.length>0){const{data}=await q;guests=(data||[]);}
    segmentLabel=`First-time + lapsed with emails · ${tplPropName}`;
  }else if(key==='upsell'){
    const{data}=await sb.rpc('get_inhouse_guests',{prop_ids:propIds});
    guests=(data||[]).filter(s=>s.nights>=3&&s.email);
    segmentLabel=`Current stayovers 3+ nights · ${tplPropName}`;
  }else if(key==='birthday'){
    // Fetch guests with DOB and calculate in JS (RPC has date math issues)
    const gIds = await getGuestIdsForProperty();
    let bdayQ = sb.from('guests')
      .select('id,full_name,email,date_of_birth,marketing_consent')
      .not('date_of_birth','is',null)
      .not('email','is',null)
      .eq('marketing_consent',true)
      ;
    if(gIds && gIds.length > 0){
      // Fetch in chunks to avoid URL limit
      let allBday = [];
      for(let i=0;i<gIds.length;i+=500){
        const chunk = gIds.slice(i,i+500);
        const{data:chunk_data} = await sb.from('guests')
          .select('id,full_name,email,date_of_birth,marketing_consent')
          .in('id', chunk)
          .not('date_of_birth','is',null)
          .not('email','is',null)
          .eq('marketing_consent',true);
        if(chunk_data) allBday = allBday.concat(chunk_data);
      }
      const today = new Date(); today.setHours(0,0,0,0);
      guests = allBday.map(g=>{
        const dob = new Date(g.date_of_birth);
        let bday = new Date(today.getFullYear(), dob.getMonth(), dob.getDate());
        if(bday < today) bday = new Date(today.getFullYear()+1, dob.getMonth(), dob.getDate());
        return {...g, days_until: Math.round((bday-today)/86400000)};
      }).filter(g=>g.days_until>=0&&g.days_until<=30)
        .sort((a,b)=>a.days_until-b.days_until);
    } else if(gIds === null){
      // All properties
      const{data:allData} = await sb.from('guests')
        .select('id,full_name,email,date_of_birth,marketing_consent')
        .not('date_of_birth','is',null)
        .not('email','is',null)
        .eq('marketing_consent',true)
        ;
      const today = new Date(); today.setHours(0,0,0,0);
      guests = (allData||[]).map(g=>{
        const dob = new Date(g.date_of_birth);
        let bday = new Date(today.getFullYear(), dob.getMonth(), dob.getDate());
        if(bday < today) bday = new Date(today.getFullYear()+1, dob.getMonth(), dob.getDate());
        return {...g, days_until: Math.round((bday-today)/86400000)};
      }).filter(g=>g.days_until>=0&&g.days_until<=30)
        .sort((a,b)=>a.days_until-b.days_until);
    }
    segmentLabel=`Birthdays next 30 days · ${tplPropName}`;
  }else if(key==='seasonal'){
    const gIds=await getGuestIdsForProperty();
    let q=sb.from('guests').select('id,full_name,email,marketing_consent').not('email','is',null).eq('marketing_consent',true);
    if(gIds){if(gIds.length===0){guests=[];}else{q=q.in('id',gIds);}}
    if(!gIds||gIds.length>0){const{data}=await q;guests=(data||[]);}
    segmentLabel=`All guests with marketing consent · ${tplPropName}`;
  }else if(key==='locals'){
    // Maldivian local guests only
    const gIds=await getGuestIdsForProperty();
    let q=sb.from('guests').select('id,full_name,email,phone,marketing_consent,nationality,guest_type')
      .eq('guest_type','local')
      .eq('marketing_consent',true);
    if(gIds){if(gIds.length===0){guests=[];}else{q=q.in('id',gIds);}}
    if(!gIds||gIds.length>0){const{data}=await q;guests=(data||[]);}
    segmentLabel=`Local Maldivian guests · ${tplPropName}`;
  }

  // Deduplicate
  const seen=new Set();const unique=[];
  guests.forEach(g=>{const id=g.id||g.guest_id;if(!seen.has(id)){seen.add(id);unique.push(g);}});
  guests=unique;

  // Build the expandable guest list
  const displayCount=5;
  const guestEmails=guests.map(g=>g.email).filter(Boolean);
  let guestListHtml=guests.slice(0,displayCount).map(g=>`<div class="row">
    <div style="flex:1;min-width:0;color:var(--gray-text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${escapeHtml(g.full_name||g.name||'')}</div>
    <div style="font-size:12px;color:var(--black)">${escapeHtml(g.email||'')}</div>
  </div>`).join('');

  if(guests.length>displayCount){
    const extraId='gl-'+Date.now();
    window['_gdata_'+extraId]=guests.map(g=>({n:g.full_name||g.name||'',e:g.email||''}));
    guestListHtml+=`<div style="font-size:12px;color:var(--kaani-orange);padding-top:8px;cursor:pointer;font-weight:500" data-id="${extraId}" onclick="expandGuestList(this)">▾ Show all ${guests.length} guests</div>`;
  }

  document.getElementById('action-content').innerHTML=`
    <div class="card">
      <div class="ch"><div class="ct">${escapeHtml(tpl.name)}</div><span class="badge be">${escapeHtml(tpl.description||'')}</span></div>
      <div class="sec">Segment: ${escapeHtml(segmentLabel)} (${guests.length} guests)</div>
      <div style="margin-bottom:14px;max-height:220px;overflow-y:auto">${guestListHtml}</div>
      <div class="sec">Subject</div>
      <div class="tpl">${escapeHtml(tpl.subject||'')}</div>
      <div class="sec">Body template</div>
      <div class="tpl" id="tpl-body-${key}">${escapeHtml(tpl.body||'')}</div>
      <div class="btns">
        <button class="btn btnp" onclick="copyToClipboard(document.getElementById('tpl-body-${key}').innerText,'Template copied')">Copy email template</button>
        <button class="btn" onclick="copyWATemplate('${key}','${escapeHtml(tplPropName)}')">📱 Copy WhatsApp</button>
        <button class="btn" onclick='copyToClipboard(${JSON.stringify(guestEmails.join(", "))},"${guestEmails.length} emails copied")'>Copy ${guestEmails.length} emails</button>
        <button class="btn" onclick="logCampaignSend('${key}',${guests.length})">Log as sent</button>
      </div>
      <div style="margin-top:10px;padding:10px 14px;background:var(--kaani-cream-deep);border-radius:8px;border:1px solid var(--line-peach)">
        <div style="font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--gray-text);font-weight:600;margin-bottom:4px">Campaign code</div>
        <div style="font-size:18px;font-weight:700;color:var(--kaani-orange);letter-spacing:.05em" id="campaign-code-${key}">${generateCampaignCode(key, STATE.selectedPropertyId||1)}</div>
        <div style="font-size:11px;color:var(--gray-text);margin-top:4px">Include this code in your email/WhatsApp. Ask guests to quote it when booking.</div>
        <button class="btn" style="margin-top:8px;font-size:11px;padding:5px 12px" onclick="copyToClipboard('${generateCampaignCode(key, STATE.selectedPropertyId||1)}','Campaign code copied')">Copy code</button>
      </div>
      <div style="margin-top:10px;font-size:12px;color:var(--gray-text)">Use <strong>{{first_name}}</strong> and <strong>{{property_name}}</strong> — replace before sending.</div>
      <div style="margin-top:12px;padding:14px;background:var(--kaani-cream-deep);border-radius:10px;border:1px solid var(--line-peach)">
        <div class="sec" style="margin-bottom:8px">WhatsApp message preview</div>
        <div id="wa-preview-${key}" style="font-size:13px;color:var(--black-soft);line-height:1.7;white-space:pre-wrap;font-family:inherit;min-height:40px;color:var(--gray-text);font-style:italic">Loading preview...</div>
      </div>
    </div>`;

  // Populate WA preview immediately after DOM update
  requestAnimationFrame(()=>{
    const wp=document.getElementById('wa-preview-'+key);
    if(wp&&WA_TEMPLATES&&WA_TEMPLATES[key]){
      wp.style.fontStyle='normal';
      wp.style.color='var(--black-soft)';
      wp.textContent=WA_TEMPLATES[key]
        .replace(/{{property_name}}/g,tplPropName)
        .replace(/{{first_name}}/g,'[Guest name]');
    }else if(wp){
      wp.textContent='WhatsApp template not configured.';
    }
  });
}



function expandGuestList(btn){
  const id=btn.getAttribute('data-id');
  const guests=window['_gdata_'+id]||[];
  const parent=btn.parentElement;
  let html=`<div style="max-height:300px;overflow-y:auto;border:1px solid var(--line-peach);border-radius:8px;margin-bottom:8px">`;
  html+=guests.map(g=>`<div style="display:flex;justify-content:space-between;align-items:center;padding:8px 12px;border-bottom:1px solid var(--line-peach);font-size:13px">
    <span style="color:var(--black-soft);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1">${escapeHtml(g.n||'')}</span>
    <span style="color:var(--gray-text);font-size:12px;margin-left:10px;flex-shrink:0">${escapeHtml(g.e||'')}</span>
  </div>`).join('');
  html+=`</div><div style="font-size:12px;color:var(--kaani-orange);cursor:pointer;font-weight:500" onclick="this.previousElementSibling.style.display='none';this.style.display='none'">▴ Collapse</div>`;
  parent.innerHTML=html;
}


async function logCampaignSend(key,count){
  const tpl=STATE.templates.find(t=>t.template_key===key);
  if(!confirm(`Log a campaign send of "${tpl.name}" to ${count} recipients?`))return;
  const{error}=await sb.from('campaigns').insert({name:tpl.name+' — '+new Date().toISOString().slice(0,10),campaign_type:key,subject:tpl.subject,body_template:tpl.body,status:'sent',sent_at:new Date().toISOString(),emails_sent:count,target_count:count,created_by:STATE.user.id});
  if(error){toast(error.message,true);return;}
  toast('Campaign logged');
}

// ============================================================================
// MARKETING
// ============================================================================
async function renderMarketingPane(){
  const el=document.getElementById('pane-marketing');
  el.innerHTML='<div class="loading">Loading</div>';

  const propIds = STATE.selectedPropertyId!==null
    ? [STATE.selectedPropertyId]
    : STATE.accessibleProperties.map(p=>p.id);

  const {data:d, error}=await sb.rpc('get_marketing_lists',{prop_ids:propIds});
  if(error){
    console.error('Marketing RPC error:',error);
    el.innerHTML='<div class="empty">Error: '+escapeHtml(error.message)+'<br>Make sure 04_rpc_functions.sql has been run in Supabase.</div>';
    return;
  }

  const allEmails   = (d?.all_emails||[]).filter(Boolean);
  const directEmails= (d?.direct_emails||[]).filter(Boolean);
  const repeatEmails= (d?.repeat_emails||[]).filter(Boolean);
  const dobEmails   = (d?.dob_emails||[]).filter(Boolean);
  // Store for clickable stat cards
  window._mktData = {allEmails, directEmails, repeatEmails, dobEmails};
  const scopeLabel  = STATE.selectedPropertyId
    ? (STATE.properties.find(p=>p.id===STATE.selectedPropertyId)?.name||'Property')
    : 'All Properties';

  el.innerHTML=`
    <div class="sg" id="mkt-stats">
      <div class="sm" style="cursor:pointer" onclick="showMktDetail('All contactable emails', window._mktData?.allEmails||[])">
        <div class="sl">Total contactable</div><div class="sv">${(d?.total_count||0).toLocaleString()}</div>
        <div class="ss" style="color:var(--kaani-orange)">tap to view ↗</div></div>
      <div class="sm" style="cursor:pointer" onclick="showMktDetail('Direct bookers', window._mktData?.directEmails||[])">
        <div class="sl">Direct bookers</div><div class="sv">${(d?.direct_count||0).toLocaleString()}</div>
        <div class="ss" style="color:var(--kaani-orange)">tap to view ↗</div></div>
      <div class="sm" style="cursor:pointer" onclick="showMktDetail('Repeat & VIP guests', window._mktData?.repeatEmails||[])">
        <div class="sl">Repeat & VIP</div><div class="sv">${(d?.repeat_count||0).toLocaleString()}</div>
        <div class="ss" style="color:var(--kaani-orange)">tap to view ↗</div></div>
      <div class="sm" style="cursor:pointer" onclick="showMktDetail('Guests with DOB', window._mktData?.dobEmails||[])">
        <div class="sl">With DOB</div><div class="sv">${(d?.dob_count||0).toLocaleString()}</div>
        <div class="ss" style="color:var(--kaani-orange)">tap to view ↗</div></div>
    </div>
    <div style="margin-bottom:12px">
      <button class="btn btnp" onclick='copyToClipboard(${JSON.stringify(allEmails.join(", "))},"All emails copied — ${allEmails.length} addresses")'>Copy all ${allEmails.length.toLocaleString()} emails</button>
    </div>
    <div id="mkt-groups"></div>`;

  // Fetch local guest emails separately
  const{data:localData} = await sb.from('stays').select('guests!inner(email,guest_type,marketing_consent)')
    .in('property_id', propIds).not('guests.email','is',null);
  const localEmails = [...new Set((localData||[])
    .map(s=>s.guests)
    .filter(g=>g&&g.guest_type==='local'&&g.marketing_consent&&g.email)
    .map(g=>g.email))];

  const groups=[
    {label:'All guests with email',detail:`${allEmails.length.toLocaleString()} contacts (consented)`,emails:allEmails},
    {label:'Direct bookings',detail:`${directEmails.length.toLocaleString()} contacts — best for loyalty offers`,emails:directEmails},
    {label:'Repeat & VIP guests',detail:`${repeatEmails.length.toLocaleString()} contacts — your most valuable`,emails:repeatEmails},
    {label:'Guests with DOB on file',detail:`${dobEmails.length.toLocaleString()} contacts — birthday campaign targets`,emails:dobEmails},
    {label:'🇲🇻 Local Maldivian guests',detail:`${localEmails.length.toLocaleString()} contacts — special fares & local offers`,emails:localEmails,highlight:true},
  ];

  document.getElementById('mkt-groups').innerHTML=groups.map(g=>`
    <div class="card" style="padding:13px 15px;display:flex;align-items:center;gap:12px;flex-wrap:wrap;${g.highlight?'border-color:var(--kaani-orange);background:var(--kaani-cream-deep)':''}">
      <div style="flex:1;min-width:0">
        <div style="font-size:13px;font-weight:600;${g.highlight?'color:var(--kaani-orange)':''}">${escapeHtml(g.label)}</div>
        <div style="font-size:11px;color:var(--gray-text);margin-top:3px">${escapeHtml(g.detail)}</div>
      </div>
      <button class="btn${g.highlight?' btnp':''}" onclick='copyToClipboard(${JSON.stringify(g.emails.join(", "))},"${g.emails.length} emails copied")'>Copy ${g.emails.length.toLocaleString()} emails</button>
    </div>`).join('');
}



function showMktDetail(title, emails){
  const sp=document.getElementById('sp');
  sp.style.display='block';
  sp.innerHTML=`<div class="panel-wrap" onclick="if(event.target.classList.contains('panel-wrap'))document.getElementById('sp').style.display='none'">
    <div class="panel" onclick="event.stopPropagation()">
      <button class="pc" onclick="document.getElementById('sp').style.display='none'">✕</button>
      <div class="pn">${escapeHtml(title)}</div>
      <div class="ps">${emails.length.toLocaleString()} guests</div>
      <button class="btn btnp" style="margin-bottom:14px;width:100%" onclick='copyToClipboard(${JSON.stringify(emails.join(", "))},"${emails.length} emails copied")'>Copy all ${emails.length.toLocaleString()} emails</button>
      <div style="max-height:65vh;overflow-y:auto">
        ${emails.map(e=>`<div style="padding:8px 0;border-bottom:1px solid var(--line-peach);font-size:13px;color:var(--black)">${escapeHtml(e)}</div>`).join('')}
      </div>
    </div>
  </div>`;
}

async function renderTemplatesPane(){
  const el=document.getElementById('pane-templates');
  const{data}=await sb.from('email_templates').select('*').order('id');
  STATE.templates=data||[];
  el.innerHTML=`
    <div class="card">
      <div class="ct" style="margin-bottom:8px">Email templates</div>
      <div class="cd">Edit your templates here — changes apply immediately to the Actions tab. Use {{first_name}} and {{property_name}} as placeholders. ${STATE.profile.role==='staff'?'<strong>Staff role: read-only.</strong>':''}</div>
    </div>
    <div id="tpl-list"></div>`;

  const canEdit=['admin','manager'].includes(STATE.profile.role);
  document.getElementById('tpl-list').innerHTML=STATE.templates.map(t=>`
    <div class="card">
      <div class="ch"><div class="ct">${escapeHtml(t.name)}</div><span class="badge bg">${escapeHtml(t.template_key)}</span></div>
      ${t.description?`<div class="cd">${escapeHtml(t.description)}</div>`:''}
      <div class="sec">Subject</div>
      <input type="text" id="tpl-sub-${t.id}" value="${escapeHtml(t.subject||'')}" ${canEdit?'':'readonly'} style="width:100%;padding:9px 12px;font-size:13px;border-radius:8px;border:1px solid #E8E6E0;font-family:inherit;margin-bottom:10px"/>
      <div class="sec">Body</div>
      <textarea id="tpl-body-${t.id}" class="tpl-edit" ${canEdit?'':'readonly'}>${escapeHtml(t.body||'')}</textarea>
      ${canEdit?`<button class="btn btnp" style="margin-top:10px" onclick="saveTemplate(${t.id})">Save changes</button>`:''}
    </div>`).join('');
}

async function saveTemplate(id){
  const subject=document.getElementById('tpl-sub-'+id).value;
  const body=document.getElementById('tpl-body-'+id).value;
  const{error}=await sb.from('email_templates').update({subject,body,updated_by:STATE.user.id}).eq('id',id);
  if(error){toast(error.message,true);return;}
  toast('Template saved');
  await loadInitialData();
}

// ============================================================================
// REPORTS — downloadable CSVs
// ============================================================================
async function renderReportsPane(){
  const el=document.getElementById('pane-reports');
  el.innerHTML=`
    <div class="card">
      <div class="ct" style="margin-bottom:8px">Reports</div>
      <div class="cd">Download CSV reports for management, NHGAM, partners, or your own analysis. Reports include all data filtered by your selected date range.</div>
    </div>
    <div class="g2">
      <div class="card">
        <div class="ct" style="margin-bottom:10px">Guest list</div>
        <div class="cd">Full guest database with contact details, lifecycle stats, and tags.</div>
        <button class="btn btnp" onclick="downloadGuestReport()">Download CSV</button>
      </div>
      <div class="card">
        <div class="ct" style="margin-bottom:10px">Stays / bookings</div>
        <div class="cd">All bookings with property, dates, channel, and revenue. Filterable by date range.</div>
        <div style="margin:10px 0;display:flex;gap:8px;align-items:center">
          <input type="date" id="rep-from" style="padding:6px 10px;border-radius:6px;border:1px solid #E8E6E0;font-family:inherit"/>
          <span style="font-size:12px;color:#5F5E5A">to</span>
          <input type="date" id="rep-to" style="padding:6px 10px;border-radius:6px;border:1px solid #E8E6E0;font-family:inherit"/>
        </div>
        <button class="btn btnp" onclick="downloadStaysReport()">Download CSV</button>
      </div>
      <div class="card">
        <div class="ct" style="margin-bottom:10px">Source channel report</div>
        <div class="cd">Booking channel breakdown — direct vs OTA vs tour operator.</div>
        <button class="btn btnp" onclick="downloadSourceReport()">Download CSV</button>
      </div>
      <div class="card">
        <div class="ct" style="margin-bottom:10px">Nationality report</div>
        <div class="cd">Guest count by nationality — useful for tourism board and partner reports.</div>
        <button class="btn btnp" onclick="downloadNationalityReport()">Download CSV</button>
      </div>
      <div class="card">
        <div class="ct" style="margin-bottom:10px">Repeat guests</div>
        <div class="cd">Guests with 2+ stays — your loyalty base.</div>
        <button class="btn btnp" onclick="downloadRepeatReport()">Download CSV</button>
      </div>
      <div class="card">
        <div class="ct" style="margin-bottom:10px">Campaign log</div>
        <div class="cd">All marketing sends — what went out, when, and to how many.</div>
        <button class="btn btnp" onclick="downloadCampaignReport()">Download CSV</button>
      </div>
    </div>`;
}

function csvEscape(v){if(v===null||v===undefined)return'';const s=String(v);if(s.includes(',')||s.includes('"')||s.includes('\n'))return'"'+s.replace(/"/g,'""')+'"';return s;}
function downloadCSV(filename,headers,rows){
  const csv=[headers.join(',')].concat(rows.map(r=>r.map(csvEscape).join(','))).join('\n');
  const blob=new Blob([csv],{type:'text/csv'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');a.href=url;a.download=filename;a.click();URL.revokeObjectURL(url);
}

async function downloadGuestReport(){
  toast('Generating guest report...');
  const propIds=getFilterPropertyIds();
  const isAll=!STATE.selectedPropertyId;
  let q=sb.from('guests').select('*').order('last_stay_date',{ascending:false});
  if(!isAll){
    const{data:propStays}=await sb.from('stays').select('guest_id').in('property_id',propIds);
    const ids=Array.from(new Set((propStays||[]).map(s=>s.guest_id)));
    if(ids.length===0){toast('No guests for this property',true);return;}
    q=q.in('id',ids);
  }
  const{data}=await q;
  const rows=(data||[]).map(g=>[g.full_name,g.email,g.phone,g.nationality,g.date_of_birth,g.passport_number,g.national_id,g.guest_type,g.lead_status,g.total_stays,g.total_nights,g.total_revenue_usd,g.first_stay_date,g.last_stay_date,g.marketing_consent]);
  const scope=STATE.selectedPropertyId?STATE.properties.find(p=>p.id===STATE.selectedPropertyId)?.name?.replace(/\s+/g,'_').toLowerCase():'all';
  downloadCSV('kaani_guests_'+scope+'_'+new Date().toISOString().slice(0,10)+'.csv',['Name','Email','Phone','Nationality','DOB','Passport','National ID','Type','Status','Stays','Nights','Revenue','First Stay','Last Stay','Marketing OK'],rows);
}

async function downloadStaysReport(){
  toast('Generating stays report...');
  const from=document.getElementById('rep-from').value;
  const to=document.getElementById('rep-to').value;
  const propIds=getFilterPropertyIds();
  let q=sb.from('stays').select('*,guests(full_name,email,nationality),properties(name)').in('property_id',propIds);
  if(from)q=q.gte('arrival_date',from);
  if(to)q=q.lte('arrival_date',to);
  const{data}=await q;
  const rows=(data||[]).map(s=>[s.guests?.full_name,s.guests?.email,s.guests?.nationality,s.properties?.name,s.arrival_date,s.departure_date,s.nights,s.room_number,s.source,s.channel_type,s.rate_per_night_usd,s.total_revenue_usd,s.status]);
  const scope=STATE.selectedPropertyId?STATE.properties.find(p=>p.id===STATE.selectedPropertyId)?.name?.replace(/\s+/g,'_').toLowerCase():'all';
  downloadCSV('kaani_stays_'+scope+'_'+new Date().toISOString().slice(0,10)+'.csv',['Guest','Email','Nationality','Property','Arrival','Departure','Nights','Room','Source','Channel','Rate','Revenue','Status'],rows);
}

async function downloadSourceReport(){
  const{data}=await sb.from('stays').select('source,channel_type,total_revenue_usd');
  const counts={};
  (data||[]).forEach(s=>{const k=s.source||'Unknown';if(!counts[k])counts[k]={count:0,revenue:0,channel:s.channel_type};counts[k].count++;counts[k].revenue+=(s.total_revenue_usd||0);});
  const rows=Object.entries(counts).sort((a,b)=>b[1].count-a[1].count).map(([s,v])=>[s,v.channel,v.count,v.revenue.toFixed(2)]);
  downloadCSV('kaani_sources_'+new Date().toISOString().slice(0,10)+'.csv',['Source','Channel Type','Stay Count','Total Revenue'],rows);
}

async function downloadNationalityReport(){
  const{data}=await sb.from('guests').select('nationality');
  const counts={};(data||[]).forEach(g=>{const k=g.nationality||'Unknown';counts[k]=(counts[k]||0)+1;});
  const rows=Object.entries(counts).sort((a,b)=>b[1]-a[1]).map(([n,v])=>[n,v]);
  downloadCSV('kaani_nationalities_'+new Date().toISOString().slice(0,10)+'.csv',['Nationality','Guest Count'],rows);
}

async function downloadRepeatReport(){
  const{data}=await sb.from('guests').select('*').gte('total_stays',2).order('total_stays',{ascending:false});
  const rows=(data||[]).map(g=>[g.full_name,g.email,g.nationality,g.total_stays,g.total_nights,g.total_revenue_usd,g.first_stay_date,g.last_stay_date]);
  downloadCSV('kaani_repeat_guests_'+new Date().toISOString().slice(0,10)+'.csv',['Name','Email','Nationality','Stays','Nights','Revenue','First Stay','Last Stay'],rows);
}

async function downloadCampaignReport(){
  const{data}=await sb.from('campaigns').select('*').order('sent_at',{ascending:false});
  const rows=(data||[]).map(c=>[c.name,c.campaign_type,c.status,c.sent_at,c.emails_sent,c.responses_received,c.bookings_attributed,c.revenue_attributed]);
  downloadCSV('kaani_campaigns_'+new Date().toISOString().slice(0,10)+'.csv',['Name','Type','Status','Sent At','Emails Sent','Responses','Bookings','Revenue'],rows);
}

// ============================================================================
// IMPORT — Ezee CSV with proper deduplication
// ============================================================================
function renderImportPane(){
  const el = document.getElementById('pane-import');
  el.innerHTML=`
    <div class="g2">
      <div class="card">
        <div class="ch"><div class="ct">Import from Ezee</div></div>
        <div class="cd">Upload one or more Ezee daily guest list CSVs at once. Guests are matched by passport / national ID — no duplicates ever created.</div>
        <div style="margin-bottom:14px">
          <label style="font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--gray-text);font-weight:600;display:block;margin-bottom:6px">Property</label>
          <select id="imp-prop" style="width:100%;padding:10px 14px;border-radius:10px;border:1px solid var(--border-subtle);font-family:inherit;background:#fff;font-size:13px">
            ${STATE.properties.map(p=>`<option value="${p.id}">${escapeHtml(p.name)}</option>`).join('')}
          </select>
        </div>

        <div id="imp-dropzone" style="border:2px dashed var(--kaani-peach);border-radius:14px;padding:40px 24px;text-align:center;background:var(--kaani-cream-deep);cursor:pointer;transition:all 0.2s"
          ondragover="event.preventDefault();this.style.borderColor='var(--kaani-orange)';this.style.background='var(--kaani-primary-soft)'"
          ondragleave="this.style.borderColor='var(--kaani-peach)';this.style.background='var(--kaani-cream-deep)'"
          ondrop="handleDrop(event)"
          onclick="document.getElementById('imp-file').click()">
          <div style="font-size:32px;margin-bottom:10px">📂</div>
          <div style="font-size:14px;font-weight:600;color:var(--black);margin-bottom:6px">Drop CSV files here</div>
          <div style="font-size:13px;color:var(--gray-text)">or <span style="color:var(--kaani-orange);text-decoration:underline;cursor:pointer">browse to upload</span></div>
          <div style="font-size:11px;color:var(--gray-muted);margin-top:8px">Multiple files supported · Ezee guest list format</div>
          <input type="file" id="imp-file" accept=".csv" multiple style="display:none" onchange="handleFileSelect(event)"/>
        </div>

        <div id="imp-file-list" style="margin-top:14px"></div>
        <div id="imp-progress-bar" style="display:none;margin-top:12px">
          <div style="display:flex;justify-content:space-between;font-size:12px;color:var(--gray-text);margin-bottom:6px">
            <span id="imp-progress-label">Processing...</span>
            <span id="imp-progress-pct">0%</span>
          </div>
          <div style="height:6px;background:var(--line-peach);border-radius:3px;overflow:hidden">
            <div id="imp-progress-fill" style="height:100%;background:var(--kaani-orange);border-radius:3px;width:0%;transition:width 0.3s"></div>
          </div>
        </div>
        <div id="imp-result" style="margin-top:12px;font-size:13px"></div>
      </div>

      <div class="card">
        <div class="ch"><div class="ct">Upload history</div><button class="btn" style="font-size:11px" onclick="loadImportHistory()">Refresh</button></div>
        <div id="imp-history"><div class="loading">Loading history</div></div>
      </div>
    </div>`;

  loadImportHistory();
}

function handleDrop(event){
  event.preventDefault();
  const dz = document.getElementById('imp-dropzone');
  dz.style.borderColor='var(--kaani-peach)';
  dz.style.background='var(--kaani-cream-deep)';
  const files = Array.from(event.dataTransfer.files).filter(f=>f.name.endsWith('.csv'));
  if(files.length===0){toast('Please drop CSV files only',true);return;}
  showFileList(files);
}

function handleFileSelect(event){
  const files = Array.from(event.target.files);
  if(files.length===0)return;
  showFileList(files);
}

function showFileList(files){
  window._pendingFiles = files;
  const el = document.getElementById('imp-file-list');
  el.innerHTML=`
    <div style="margin-bottom:10px">
      <div class="sec">${files.length} file${files.length>1?'s':''} selected</div>
      ${files.map((f,i)=>`<div style="display:flex;align-items:center;gap:10px;padding:8px 12px;border-radius:8px;background:var(--kaani-cream-deep);border:1px solid var(--line-peach);margin-bottom:6px;font-size:13px">
        <span style="font-size:16px">📄</span>
        <div style="flex:1;min-width:0">
          <div style="font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${escapeHtml(f.name)}</div>
          <div style="font-size:11px;color:var(--gray-text)">${(f.size/1024).toFixed(1)} KB</div>
        </div>
        <div id="file-status-${i}" style="font-size:11px;color:var(--gray-muted)">Pending</div>
      </div>`).join('')}
    </div>
    <button class="btn btnp" style="width:100%" onclick="startImport()">Import ${files.length} file${files.length>1?'s':''}</button>`;
}

async function startImport(){
  const files = window._pendingFiles;
  if(!files||files.length===0){toast('No files selected',true);return;}
  const propId = parseInt(document.getElementById('imp-prop').value);
  const propName = STATE.properties.find(p=>p.id===propId)?.name||'Unknown';

  const progBar = document.getElementById('imp-progress-bar');
  const progLabel = document.getElementById('imp-progress-label');
  const progPct = document.getElementById('imp-progress-pct');
  const progFill = document.getElementById('imp-progress-fill');
  progBar.style.display='block';

  let totalAdded=0, totalUpdated=0, totalErrors=0;

  for(let fi=0; fi<files.length; fi++){
    const file = files[fi];
    const statusEl = document.getElementById('file-status-'+fi);
    if(statusEl) statusEl.innerHTML='<span style="color:var(--kaani-orange)">Processing...</span>';

    progLabel.textContent = `Processing ${fi+1} of ${files.length}: ${file.name}`;
    progPct.textContent = Math.round((fi/files.length)*100)+'%';
    progFill.style.width = Math.round((fi/files.length)*100)+'%';

    try {
      const text = await file.text();
      const result = await importCSV(text, propId);
      totalAdded += result.added;
      totalUpdated += result.updated;
      totalErrors += result.errors;

      if(statusEl) statusEl.innerHTML=`<span style="color:var(--success)">✓ ${result.added} new · ${result.updated} updated</span>`;
    } catch(err) {
      totalErrors++;
      if(statusEl) statusEl.innerHTML=`<span style="color:var(--danger)">✗ Error: ${err.message}</span>`;
    }
  }

  progPct.textContent = '100%';
  progFill.style.width = '100%';
  progLabel.textContent = 'Complete!';

  // Log the import
  const importLog = {
    name: `Ezee import · ${propName} · ${new Date().toLocaleDateString('en-GB')}`,
    campaign_type: 'custom',
    notes: `Files: ${files.length} · New guests: ${totalAdded} · Updated: ${totalUpdated} · Errors: ${totalErrors}`,
    status: 'sent',
    sent_at: new Date().toISOString(),
    created_at: new Date().toISOString(),
    emails_sent: totalAdded + totalUpdated,
    target_count: totalAdded + totalUpdated,
    created_by: STATE.user.id
  };
  const{error:logErr} = await sb.from('campaigns').insert(importLog);
  if(logErr) console.error('Import log error:', logErr.message);
  else console.log('Import logged successfully');

  const result = document.getElementById('imp-result');
  result.innerHTML=`<div style="padding:12px 14px;background:var(--success-bg);border-radius:10px;color:var(--success);font-weight:500">
    ✓ Import complete — ${totalAdded} new guests · ${totalUpdated} updated · ${totalErrors} errors
  </div>`;

  toast(`Done! ${totalAdded} new guests, ${totalUpdated} updated`);
  await loadInitialData();
  loadImportHistory();
  window._pendingFiles = null;
}

async function loadImportHistory(){
  const el = document.getElementById('imp-history');
  if(!el) return;
  const{data, error:histErr} = await sb.from('campaigns').select('*').order('created_at',{ascending:false,nullsFirst:false});
  if(histErr){console.error('History error:',histErr.message);}
  if(!data||data.length===0){
    el.innerHTML='<div class="empty">No imports yet</div>';
    return;
  }
  el.innerHTML = data.map(c=>`
    <div style="padding:10px 0;border-bottom:1px solid var(--line-peach);font-size:13px">
      <div style="display:flex;align-items:flex-start;justify-content:space-between;gap:8px">
        <div style="flex:1;min-width:0">
          <div style="font-weight:500;color:var(--black)">${escapeHtml(c.name)}</div>
          <div style="font-size:11px;color:var(--gray-text);margin-top:3px">${escapeHtml(c.notes||'')}</div>
        </div>
        <div style="font-size:11px;color:var(--gray-muted);white-space:nowrap">${c.sent_at?new Date(c.sent_at).toLocaleDateString('en-GB',{day:'numeric',month:'short',year:'numeric'}):''}</div>
      </div>
    </div>`).join('');
}



// ============================================================================
// CSV IMPORT ENGINE
// ============================================================================
function parseCSV(text){
  if(text.charCodeAt(0)===0xFEFF)text=text.slice(1);
  const lines=text.replace(/\r/g,'').split('\n').filter(l=>l.trim());
  if(lines.length<2)return{headers:[],rows:[]};
  function parseLine(line){
    const result=[];let cur='';let inQ=false;
    for(let i=0;i<line.length;i++){
      const c=line[i];
      if(c==='"'){if(inQ&&line[i+1]==='"'){cur+='"';i++;}else{inQ=!inQ;}}
      else if(c===','&&!inQ){result.push(cur);cur='';}
      else cur+=c;
    }
    result.push(cur);
    return result;
  }
  const headers=parseLine(lines[0]).map(h=>h.replace(/^\uFEFF/,'').replace(/^"|"$/g,'').trim());
  const rows=lines.slice(1).map(parseLine);
  return{headers,rows};
}

function extractDOB(addr){
  if(!addr||!addr.trim()||addr.trim().toLowerCase()==='maldives.')return null;
  const cleaned=addr.replace(/^(?:dob|DOB)[:\s]*/i,'').trim();
  let m=cleaned.match(/(\d{1,2})[.\-](\d{1,2})[.\-](\d{4})/);
  if(m){const[,d,mo,y]=m;return`${y}-${mo.padStart(2,'0')}-${d.padStart(2,'0')}`;}
  m=cleaned.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if(m){const a=+m[1],b=+m[2],y=+m[3];if(a>12)return`${y}-${String(b).padStart(2,'0')}-${String(a).padStart(2,'0')}`;if(b>12)return`${y}-${String(a).padStart(2,'0')}-${String(b).padStart(2,'0')}`;return`${y}-${String(b).padStart(2,'0')}-${String(a).padStart(2,'0')}`;}
  return null;
}

function parseEzeeDate(d){
  if(!d)return null;
  let m=d.match(/^(\d{1,2})-(\d{1,2})-(\d{2,4})$/);
  if(m){const[,dd,mm,yy]=m;const y=yy.length===2?'20'+yy:yy;return`${y}-${mm.padStart(2,'0')}-${dd.padStart(2,'0')}`;}
  return null;
}

function classifyChannel(source){
  if(!source)return'other';
  const s=source.toUpperCase();
  if(s==='DIRECT BOOKING'||s.includes('DIRECT'))return'direct';
  if(s.includes('BOOKING.COM')||s.includes('AGODA')||s.includes('EXPEDIA')||s.includes('MAKE MY TRIP')||s.includes('HOTELS.COM'))return'ota';
  if(s.includes('DESTINOS')||s.includes('TOUR'))return'tour_operator';
  if(s.includes('WALK'))return'walk_in';
  return'other';
}

// Known placeholder emails - treated as no real email

async function importCSV(text, propId){
  const{headers,rows}=parseCSV(text);
  const idx={};headers.forEach((h,i)=>idx[h]=i);

  const required=['Identity','Guest Name','Arrival','Departure','Nights'];
  for(const r of required){
    if(!(r in idx))throw new Error('Missing column: '+r+'. Found: '+headers.join(', '));
  }

  let added=0,updated=0,errors=0;

  for(let i=0;i<rows.length;i++){
    const r=rows[i];
    const identity=r[idx['Identity']]?.trim();
    if(!identity)continue;

    const isNationalId=identity.toLowerCase().includes('national id');
    const passport=isNationalId?null:identity.replace(/^Passport-/i,'');
    const nationalId=isNationalId?identity.replace(/^National\s*ID-/i,''):null;

    const name=r[idx['Guest Name']]?.trim()||'';
    let email=r[idx['Email']]?.trim()||null;
    if(email&&isPlaceholderEmail(email))email=null;
    const nationality=r[idx['Nationality']]?.trim()||null;
    const address=r[idx['Address']]?.trim()||null;
    const dob=extractDOB(address);
    const nameParts=name.replace(/^(Mr\.|Ms\.|Mrs\.|Dr\.)\s+/i,'').trim().split(/\s+/);
    const firstName=nameParts[0]||null;
    const lastName=nameParts.length>1?nameParts[nameParts.length-1]:null;

    try {
      // Find or create guest
      let guestId;
      let existing=null;
      if(passport){const{data}=await sb.from('guests').select('id').eq('passport_number',passport).maybeSingle();existing=data;}
      if(!existing&&nationalId){const{data}=await sb.from('guests').select('id').eq('national_id',nationalId).maybeSingle();existing=data;}
      if(!existing&&email&&!isPlaceholderEmail(email)){const{data}=await sb.from('guests').select('id').ilike('email',email).maybeSingle();existing=data;}

      if(existing){
        guestId=existing.id;
        const updates={full_name:name,first_name:firstName,last_name:lastName};
        if(email)updates.email=email;
        if(nationality)updates.nationality=nationality;
        if(dob)updates.date_of_birth=dob;
        if(passport)updates.passport_number=passport;
        if(nationalId)updates.national_id=nationalId;
        await sb.from('guests').update(updates).eq('id',guestId);
        updated++;
      } else {
        const guestType=r[idx['Guest Type']]?.includes('Local')?'local':'tourist';
        const{data,error}=await sb.from('guests').insert({full_name:name,first_name:firstName,last_name:lastName,email,nationality,date_of_birth:dob,passport_number:passport,national_id:nationalId,guest_type:guestType,created_by:STATE.user.id}).select('id').single();
        if(error){errors++;continue;}
        guestId=data.id;
        added++;
      }

      // Add stay
      const arrival=parseEzeeDate(r[idx['Arrival']]);
      const departure=parseEzeeDate(r[idx['Departure']]);
      const nights=parseInt(r[idx['Nights']])||0;
      const source=r[idx['Source']]?.trim()||null;
      const status=(r[idx['Status']]?.toLowerCase().replace(/\s+/g,'_'))||'reservation';
      const paxStr=r[idx['Pax']]||'';
      const paxMatch=paxStr.match(/(\d+)\s*\(A\)\s*\/\s*(\d+)\s*\(C\)/);
      const paxA=paxMatch?+paxMatch[1]:1;
      const paxC=paxMatch?+paxMatch[2]:0;

      const{data:existingStay}=await sb.from('stays').select('id').eq('guest_id',guestId).eq('property_id',propId).eq('arrival_date',arrival).maybeSingle();

      const stayData={guest_id:guestId,property_id:propId,arrival_date:arrival,departure_date:departure,nights,room_number:r[idx['Room']]?.trim()||null,rate_type:r[idx['Rate Type']]?.trim()||null,pax_adults:paxA,pax_children:paxC,source,channel_type:classifyChannel(source),status,reservation_remark:r[idx['Reservation Remark']]?.trim()||null};

      if(existingStay){
        await sb.from('stays').update(stayData).eq('id',existingStay.id);
      } else {
        await sb.from('stays').insert(stayData);
      }
    } catch(err){
      errors++;
    }
  }

  return{added,updated,errors};
}

// ============================================================================
// ADMIN
// ============================================================================

async function dismissDuplicate(id1, id2, btn){
  await sb.from('communications').insert([
    {guest_id:id1,channel:'note',direction:'internal',body:`Confirmed different person — dismissed duplicate match with guest ${id2}`,staff_user_id:STATE.user.id},
    {guest_id:id2,channel:'note',direction:'internal',body:`Confirmed different person — dismissed duplicate match with guest ${id1}`,staff_user_id:STATE.user.id}
  ]).catch(()=>{});
  const row = btn.closest('div[style*="padding:14px"]');
  if(row){row.style.opacity='0.3';row.style.pointerEvents='none';}
  toast('Marked as different people');
}

async function openMergePanel(id1, name1, id2, name2){
  // Load full details of both guests
  const [r1, r2, stays1, stays2] = await Promise.all([
    sb.from('guests').select('*').eq('id',id1).single(),
    sb.from('guests').select('*').eq('id',id2).single(),
    sb.from('stays').select('*,properties(name)').eq('guest_id',id1).order('arrival_date',{ascending:false}),
    sb.from('stays').select('*,properties(name)').eq('guest_id',id2).order('arrival_date',{ascending:false})
  ]);

  const g1 = r1.data || {}; const g2 = r2.data || {};
  const s1 = stays1.data || []; const s2 = stays2.data || [];

  const sp = document.getElementById('sp');
  sp.style.display = 'block';
  sp.innerHTML = `<div class="panel-wrap" onclick="if(event.target.classList.contains('panel-wrap'))document.getElementById('sp').style.display='none'">
    <div class="panel" onclick="event.stopPropagation()" style="width:520px">
      <button class="pc" onclick="document.getElementById('sp').style.display='none'">✕</button>
      <div class="pn">Merge guests</div>
      <div class="ps">Choose which record to keep as the primary. All stays from the other record will be moved across, then the duplicate is deleted.</div>

      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:16px 0">
        <div style="padding:14px;border:2px solid var(--line-peach);border-radius:12px;cursor:pointer" id="pick-1" onclick="selectMergePrimary(1)">
          <div style="font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--gray-text);margin-bottom:8px">Option A</div>
          <div style="font-weight:700;font-size:14px">${escapeHtml(g1.full_name||'')}</div>
          <div style="font-size:12px;color:var(--gray-text);margin-top:4px">${escapeHtml(g1.email||'no email')}</div>
          <div style="font-size:12px;color:var(--gray-text)">${escapeHtml(g1.nationality||'')} · ${escapeHtml(g1.passport_number||g1.national_id||'no ID')}</div>
          <div style="font-size:12px;color:var(--gray-text)">${s1.length} stay(s)</div>
          <div style="margin-top:8px;font-size:11px;color:var(--kaani-orange);font-weight:600">Click to keep this one</div>
        </div>
        <div style="padding:14px;border:2px solid var(--line-peach);border-radius:12px;cursor:pointer" id="pick-2" onclick="selectMergePrimary(2)">
          <div style="font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--gray-text);margin-bottom:8px">Option B</div>
          <div style="font-weight:700;font-size:14px">${escapeHtml(g2.full_name||'')}</div>
          <div style="font-size:12px;color:var(--gray-text);margin-top:4px">${escapeHtml(g2.email||'no email')}</div>
          <div style="font-size:12px;color:var(--gray-text)">${escapeHtml(g2.nationality||'')} · ${escapeHtml(g2.passport_number||g2.national_id||'no ID')}</div>
          <div style="font-size:12px;color:var(--gray-text)">${s2.length} stay(s)</div>
          <div style="margin-top:8px;font-size:11px;color:var(--kaani-orange);font-weight:600">Click to keep this one</div>
        </div>
      </div>

      <div id="merge-confirm" style="display:none;padding:14px;background:var(--kaani-cream-deep);border-radius:10px;border:1px solid var(--kaani-peach);margin-bottom:14px;font-size:13px"></div>

      <button class="btn btnp" id="merge-btn" style="width:100%;display:none" onclick="executeMerge('${id1}','${id2}')">Confirm merge</button>
      <div style="font-size:12px;color:var(--gray-text);margin-top:8px;text-align:center">This cannot be undone. Back up your data first if unsure.</div>
    </div>
  </div>`;

  window._mergeData = {id1, name1, id2, name2, g1, g2, primaryId: null};
}

function selectMergePrimary(which){
  const d = window._mergeData;
  if(!d) return;
  d.primaryId = which === 1 ? d.id1 : d.id2;
  d.deleteId = which === 1 ? d.id2 : d.id1;
  d.primaryName = which === 1 ? d.name1 : d.name2;
  d.deleteName = which === 1 ? d.name2 : d.name1;

  document.getElementById('pick-1').style.borderColor = which===1 ? 'var(--kaani-orange)' : 'var(--line-peach)';
  document.getElementById('pick-2').style.borderColor = which===2 ? 'var(--kaani-orange)' : 'var(--line-peach)';

  const conf = document.getElementById('merge-confirm');
  conf.style.display = 'block';
  conf.innerHTML = `<strong>Keep:</strong> ${escapeHtml(d.primaryName)}<br><strong>Delete:</strong> ${escapeHtml(d.deleteName)} (all their stays will move to the kept record)`;

  document.getElementById('merge-btn').style.display = 'block';
}

async function executeMerge(id1, id2){
  const d = window._mergeData;
  if(!d || !d.primaryId) return;
  
  const btn = document.getElementById('merge-btn');
  btn.textContent = 'Merging...'; btn.disabled = true;

  try {
    // Move all stays from deleted guest to primary
    await sb.from('stays').update({guest_id: d.primaryId}).eq('guest_id', d.deleteId);
    // Move communications
    await sb.from('communications').update({guest_id: d.primaryId}).eq('guest_id', d.deleteId);
    // Move tags (ignore conflicts)
    await sb.from('guest_tags').update({guest_id: d.primaryId}).eq('guest_id', d.deleteId).catch(()=>{});
    // Delete the duplicate guest
    await sb.from('guests').delete().eq('id', d.deleteId);
    // Recalculate primary guest totals
    const{data:stays}=await sb.from('stays').select('nights,total_revenue_usd').eq('guest_id',d.primaryId);
    const totalNights = (stays||[]).reduce((a,s)=>a+(s.nights||0),0);
    const totalRevenue = (stays||[]).reduce((a,s)=>a+(parseFloat(s.total_revenue_usd)||0),0);
    const totalStays = (stays||[]).length;
    await sb.from('guests').update({
      total_stays: totalStays,
      total_nights: totalNights,
      total_revenue_usd: totalRevenue,
      updated_at: new Date().toISOString()
    }).eq('id', d.primaryId);

    document.getElementById('sp').style.display='none';
    toast('Guests merged successfully');
    await loadInitialData();
    renderAdminPane();
  } catch(err) {
    btn.textContent = 'Confirm merge'; btn.disabled = false;
    toast('Merge failed: '+err.message, true);
  }
}

async function renderAdminPane(){
  const pane = document.getElementById('pane-admin');
  if(!pane) return;
  if(!STATE.profile || STATE.profile.role !== 'admin'){
    pane.innerHTML='<div class="empty">Admin access required</div>';
    return;
  }
  pane.innerHTML='<div class="loading">Loading admin panel</div>';

  try {
    const[usersRes, dupsRes] = await Promise.all([
      sb.from('user_profiles').select('*').order('created_at'),
      sb.from('v_potential_duplicates').select('*').limit(200).then(r=>r.error?{data:[]}:r)
    ]);

    if(usersRes.error){
      pane.innerHTML='<div class="empty">Error loading users: '+escapeHtml(usersRes.error.message)+'</div>';
      return;
    }

    const users = usersRes.data || [];
    const dups = dupsRes.data || [];

    pane.innerHTML=`
      <div class="card">
        <div class="ch"><div class="ct">Team members</div></div>
        <div class="cd" style="margin-bottom:12px">Add team members via Supabase → Authentication → Users → Add user (tick Auto Confirm). Assign roles and property access below.</div>
        ${users.map(u=>`<div class="row">
          <div style="flex:1;min-width:0">
            <div class="rn">${escapeHtml(u.full_name||u.user_id)}</div>
            <div style="font-size:11px;color:var(--gray-text)">${u.role}${u.user_id===STATE.user.id?' · (you)':''}</div>
          </div>
          <select onchange="updateUserRole('${u.user_id}',this.value)" ${u.user_id===STATE.user.id?'disabled':''} style="font-size:12px;padding:7px 10px;border-radius:8px;border:1px solid var(--line-peach);font-family:inherit;background:#fff;margin-right:8px">
            <option value="staff" ${u.role==='staff'?'selected':''}>Staff</option>
            <option value="manager" ${u.role==='manager'?'selected':''}>Manager</option>
            <option value="admin" ${u.role==='admin'?'selected':''}>Admin</option>
          </select>
          <button class="btn" style="font-size:11px;padding:6px 12px" onclick="manageProperties('${u.user_id}','${escapeHtml(u.full_name||'')}')">Properties</button>
        </div>`).join('')}
      </div>

      <div class="card">
        <div class="ch">
          <div class="ct">Potential duplicates (${dups.length})</div>
          <div class="cs">Review each pair — merge or dismiss</div>
        </div>
        ${dups.length === 0
          ? '<div class="empty">No duplicates detected — your data is clean ✓</div>'
          : dups.map(d=>`<div style="padding:14px 0;border-bottom:1px solid var(--line-peach)">
              <div style="display:flex;align-items:flex-start;gap:12px;flex-wrap:wrap">
                <div style="flex:1;min-width:160px">
                  <div style="font-size:13px;font-weight:600">${escapeHtml(d.name_1)}</div>
                  <div style="font-size:11px;color:var(--gray-text);margin-top:2px">${escapeHtml(d.email_1||'no email')}</div>
                </div>
                <div style="font-size:11px;color:var(--kaani-orange);font-weight:700;padding-top:2px;white-space:nowrap">↔ ${escapeHtml(d.match_reason?.replace('_',' ')||'')}</div>
                <div style="flex:1;min-width:160px;text-align:right">
                  <div style="font-size:13px;font-weight:600">${escapeHtml(d.name_2)}</div>
                  <div style="font-size:11px;color:var(--gray-text);margin-top:2px">${escapeHtml(d.email_2||'no email')}</div>
                </div>
              </div>
              <div style="display:flex;gap:8px;margin-top:10px;justify-content:flex-end">
                <button class="btn btnp" style="font-size:11px;padding:6px 12px" onclick="openMergePanel('${d.guest_1_id}','${escapeHtml(d.name_1)}','${d.guest_2_id}','${escapeHtml(d.name_2)}')">⚡ Merge</button>
                <button class="btn" style="font-size:11px;padding:6px 12px" onclick="dismissDuplicate('${d.guest_1_id}','${d.guest_2_id}',this)">✕ Not a duplicate</button>
              </div>
            </div>`).join('')
        }
      </div>

      <div class="card">
        <div class="ct" style="margin-bottom:12px">Quick backup</div>
        <div class="btns">
          <button class="btn" onclick="downloadGuestReport()">Download all guests (CSV)</button>
          <button class="btn" onclick="downloadStaysReport()">Download all stays (CSV)</button>
        </div>
      </div>`;
  } catch(err) {
    console.error('Admin render error:', err);
    pane.innerHTML='<div class="empty">Admin error: '+escapeHtml(err.message||'Unknown error')+'</div>';
  }
}


async function updateUserRole(userId,role){
  const{error}=await sb.from('user_profiles').update({role}).eq('user_id',userId);
  if(error){toast(error.message,true);return;}
  toast('Role updated');
}

async function manageProperties(userId,userName){
  const{data:profile}=await sb.from('user_profiles').select('property_ids').eq('user_id',userId).single();
  const current=profile?.property_ids||[];
  const sp=document.getElementById('sp');
  sp.style.display='block';
  sp.innerHTML=`<div class="panel-wrap" onclick="closePropPanel(event)">
    <div class="panel" onclick="event.stopPropagation()">
      <button class="pc" onclick="closePropPanelDirect()">✕</button>
      <div class="pn">Property access</div>
      <div class="ps">${escapeHtml(userName||'team member')}</div>
      <div class="cd">Select which properties this user can access. Leave all unchecked to grant access to all properties.</div>
      ${STATE.properties.map(p=>`<label style="display:flex;align-items:center;gap:12px;padding:12px 14px;border:1px solid var(--line-peach);border-radius:10px;margin-bottom:8px;cursor:pointer;background:var(--kaani-cream)">
        <input type="checkbox" id="prop-chk-${p.id}" ${current.includes(p.id)?'checked':''} style="width:16px;height:16px;accent-color:var(--kaani-orange);cursor:pointer">
        <span style="font-weight:500">${escapeHtml(p.name)}</span>
        <span style="margin-left:auto;font-size:11px;color:var(--gray-text)">${escapeHtml(p.code)}</span>
      </label>`).join('')}
      <button class="btn btnp" style="width:100%;margin-top:14px" onclick="savePropertyAccess('${userId}')">Save access</button>
    </div>
  </div>`;
}

function closePropPanel(e){if(e&&!e.target.classList.contains('panel-wrap'))return;closePropPanelDirect();}
function closePropPanelDirect(){document.getElementById('sp').style.display='none';}

async function savePropertyAccess(userId){
  const selected=STATE.properties.map(p=>p.id).filter(id=>{const cb=document.getElementById('prop-chk-'+id);return cb&&cb.checked;});
  const{error}=await sb.from('user_profiles').update({property_ids:selected.length>0?selected:null}).eq('user_id',userId);
  if(error){toast(error.message,true);return;}
  toast('Property access updated');
  closePropPanelDirect();
  renderAdminPane();
}

// ============================================================================
// INIT
// ============================================================================
checkSession();
document.getElementById('login-password').addEventListener('keypress',e=>{if(e.key==='Enter')doLogin();});

async function changePassword(){
  const newPwd = prompt('Enter new password (min 6 characters):');
  if(!newPwd||newPwd.length<6){toast('Password must be at least 6 characters',true);return;}
  const{error}=await sb.auth.updateUser({password:newPwd});
  if(error){toast(error.message,true);return;}
  toast('Password updated successfully');
}

