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
  selectedPropertyId: 'all',  // 'all' or a property ID number
};

const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const AVCOLS = ['#E1F5EE/#0F6E56','#E6F1FB/#185FA5','#FAEEDA/#633806','#FBEAF0/#72243E','#EAF3DE/#27500A','#EEEDFE/#3C3489'];

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
  document.getElementById('user-display').textContent=`${STATE.profile.full_name||STATE.user.email} (${STATE.profile.role})`;
  if(STATE.profile.role==='admin'){
    document.getElementById('admin-tab').style.display='inline-block';
  }
  await loadInitialData();
  setupPropertyAccess();
  renderPropertySelector();
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

  if(STATE.accessibleProperties.length<=1 && STATE.profile.role!=='admin'){
    bar.style.display='none';
    return;
  }

  bar.style.display='flex';

  let html='';
  if(STATE.profile.role==='admin'||(STATE.profile.role==='manager'&&STATE.accessibleProperties.length>1)){
    const allLabel=STATE.profile.role==='admin'?'All Properties':'All My Properties';
    html+=`<button class="pp ${STATE.selectedPropertyId===null?'on':''}" onclick="selectProperty(null)">${allLabel}</button>`;
  }
  STATE.accessibleProperties.forEach(p=>{
    html+=`<button class="pp ${STATE.selectedPropertyId===p.id?'on':''}" onclick="selectProperty(${p.id})">${escapeHtml(p.name)}</button>`;
  });

  pills.innerHTML=html;
}

async function selectProperty(propId){
  STATE.selectedPropertyId=propId;
  STATE.guestsPage=0;
  STATE.chartsRendered=false;
  renderPropertySelector();

  const activePane=document.querySelector('.pane.on');
  if(!activePane)return;
  const tabId=activePane.id.replace('pane-','');
  if(tabId==='overview')renderOverview();
  if(tabId==='dashboard')renderDashboard();
  if(tabId==='guests')renderGuestsPane();
  if(tabId==='actions')renderActionsPane();
  if(tabId==='marketing')renderMarketingPane();
  if(tabId==='reports')renderReportsPane();

  toast(propId?`Viewing: ${STATE.properties.find(p=>p.id===propId)?.name||''}`:'Viewing: All Properties');
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
function renderPropertySelector() {
  const bar = document.getElementById('property-bar');
  if (!bar) return;
  const sel = STATE.selectedPropertyId;
  const all = sel === 'all';
  const selectedName = all ? 'All Properties' : (STATE.properties.find(p => p.id == sel)?.name || 'Unknown');
  let html = `<div class="prop-selector">
    <div class="prop-label">Viewing:</div>
    <button class="prop-pill ${all ? 'on' : ''}" onclick="changeProperty('all')">All Properties</button>`;
  STATE.properties.forEach(p => {
    html += `<button class="prop-pill ${sel == p.id ? 'on' : ''}" onclick="changeProperty(${p.id})">${escapeHtml(p.name)}</button>`;
  });
  html += `</div>`;
  bar.innerHTML = html;
}

function changeProperty(id) {
  STATE.selectedPropertyId = id;
  renderPropertySelector();
  // Re-render currently active tab
  const activeBtn = document.querySelector('.nav .nb.on');
  if (activeBtn) {
    const onclick = activeBtn.getAttribute('onclick');
    const match = onclick.match(/sw\('(\w+)'/);
    if (match) {
      const tab = match[1];
      if (tab === 'overview') renderOverview();
      if (tab === 'dashboard') { STATE.chartsRendered = false; renderDashboard(); }
      if (tab === 'guests') { STATE.guestsPage = 0; renderGuestsPane(); }
      if (tab === 'actions') renderActionsPane();
      if (tab === 'marketing') renderMarketingPane();
      if (tab === 'reports') renderReportsPane();
    }
  }
}

// Helper: applies property filter to a Supabase query on the stays table
function applyPropertyFilter(query) {
  if (STATE.selectedPropertyId === 'all') return query;
  return query.eq('property_id', STATE.selectedPropertyId);
}

// Helper: get list of guest IDs that have stayed at the selected property
async function getGuestIdsForProperty() {
  if (STATE.selectedPropertyId === 'all') return null;  // null = all guests
  const { data } = await sb.from('stays').select('guest_id').eq('property_id', STATE.selectedPropertyId);
  return [...new Set((data || []).map(s => s.guest_id))];
}

// ============================================================================
// OVERVIEW
// ============================================================================
async function renderOverview(){
  const el=document.getElementById('pane-overview');
  el.innerHTML='<div class="loading">Loading</div>';

  let inhouseQuery = sb.from('stays').select('*,guests(*),properties(name)').in('status',['checked_in','stayover']).order('arrival_date',{ascending:false}).limit(8);
  let coQuery = sb.from('stays').select('*,guests(*)').eq('status','checked_out').gte('departure_date',new Date(Date.now()-3*864e5).toISOString().slice(0,10)).order('departure_date',{ascending:false}).limit(5);
  if(STATE.selectedPropertyId !== 'all'){
    inhouseQuery = inhouseQuery.eq('property_id', STATE.selectedPropertyId);
    coQuery = coQuery.eq('property_id', STATE.selectedPropertyId);
  }
  const guestIdsForProp = await getGuestIdsForProperty();
  const[statsRes,inhouseRes,bdayRes,recentCheckoutRes]=await Promise.all([
    computePropertyStats(),
    inhouseQuery,
    Promise.resolve(null),
    coQuery
  ]);

  async function computePropertyStats(){
    if(STATE.selectedPropertyId === 'all'){
      return await sb.from('v_dashboard_stats').select('*').single();
    }
    // Compute stats for the specific property
    const propStays = await sb.from('stays').select('*,guests(email,date_of_birth)').eq('property_id', STATE.selectedPropertyId);
    const stays = propStays.data || [];
    const uniqueGuests = new Set(stays.map(s => s.guest_id));
    const inHouse = stays.filter(s => ['checked_in','stayover'].includes(s.status)).length;
    const direct = stays.filter(s => s.channel_type === 'direct').length;
    const repeat = await sb.from('guests').select('id').in('id', [...uniqueGuests]).in('lead_status',['repeat','vip']);
    const totalRev = stays.reduce((a,s) => a + (parseFloat(s.net_revenue_usd) || 0), 0);
    return {data: {
      total_guests: uniqueGuests.size,
      in_house: inHouse,
      repeat_guests: (repeat.data || []).length,
      direct_bookings: direct,
      total_stays: stays.length,
      total_net_revenue: totalRev
    }};
  }

  // Build stats from property-filtered data
  const propStays=stayStatsRes.data||[];
  const guestIds=new Set(propStays.map(s=>s.guest_id));
  const allGuests=guestStatsRes.data||[];
  const propGuests=allGuests.filter(g=>guestIds.has(g.id));

  const stats={
    total_guests:propGuests.length,
    in_house:propStays.filter(s=>['checked_in','stayover'].includes(s.status)).length,
    repeat_guests:propGuests.filter(g=>['repeat','vip'].includes(g.lead_status)).length,
    direct_bookings:propStays.filter(s=>s.channel_type==='direct').length,
    total_stays:propStays.length,
    total_net_revenue:propStays.reduce((a,s)=>a+(parseFloat(s.net_revenue_usd)||0),0)
  };

  // Birthday fallback if RPC doesn't exist
  let birthdays=[];
  if(bdayRes&&bdayRes.data){
    birthdays=bdayRes.data;
  }else{
    let bdQ = sb.from('guests').select('id,full_name,email,date_of_birth').not('date_of_birth','is',null).not('email','is',null).limit(500);
    if(guestIdsForProp){
      if(guestIdsForProp.length === 0){ birthdays = []; }
      else { bdQ = bdQ.in('id', guestIdsForProp); }
    }
    if(!guestIdsForProp || guestIdsForProp.length > 0){
      const{data:bd}=await bdQ;
      birthdays=(bd||[]).filter(g=>{const d=bdayDays(g.date_of_birth);return d!==null&&d<=30;}).sort((a,b)=>bdayDays(a.date_of_birth)-bdayDays(b.date_of_birth));
    }
  }

  let html=`
    <div class="sg">
      <div class="sm"><div class="sl">Total guests</div><div class="sv">${(stats.total_guests||0).toLocaleString()}</div></div>
      <div class="sm"><div class="sl">In-house</div><div class="sv">${stats.in_house||0}</div></div>
      <div class="sm"><div class="sl">Repeat guests</div><div class="sv">${stats.repeat_guests||0}</div></div>
      <div class="sm"><div class="sl">Direct bookings</div><div class="sv">${stats.direct_bookings||0}</div></div>
      <div class="sm"><div class="sl">Total stays</div><div class="sv">${(stats.total_stays||0).toLocaleString()}</div></div>
      <div class="sm"><div class="sl">Net revenue</div><div class="sv">$${Math.round(stats.total_net_revenue||0).toLocaleString()}</div></div>
    </div>
    <div class="g2">
      <div class="card">
        <div class="ct" style="margin-bottom:12px">Today's urgent actions</div>
        <div id="ov-alerts"></div>
      </div>
      <div class="card">
        <div class="ct" style="margin-bottom:12px">Upcoming birthdays (30 days)</div>
        <div id="ov-bdays"></div>
      </div>
    </div>
    <div class="card">
      <div class="ct" style="margin-bottom:12px">Current in-house guests</div>
      <div id="ov-inhouse"></div>
    </div>`;
  el.innerHTML=html;

  // Alerts
  const recentCo=recentCheckoutRes.data||[];
  const inhouse=inhouseRes.data||[];
  const items=[];
  recentCo.slice(0,3).forEach(s=>{if(s.guests?.email)items.push({dot:'#A32D2D',txt:fn(s.guests.full_name)+' — send review request',sub:`Checked out ${fmtDate(s.departure_date)}`});});
  inhouse.filter(s=>s.nights>=3&&s.guests?.email).slice(0,2).forEach(s=>items.push({dot:'#1D9E75',txt:fn(s.guests.full_name)+' — upsell excursions',sub:`Stayover until ${fmtDate(s.departure_date)}`}));
  document.getElementById('ov-alerts').innerHTML=items.length?items.map(it=>`<div class="row"><div class="dot" style="background:${it.dot}"></div><div style="flex:1;min-width:0"><div class="rn">${escapeHtml(it.txt)}</div><div style="font-size:11px;color:#5F5E5A">${escapeHtml(it.sub)}</div></div></div>`).join(''):'<div class="empty">No urgent actions.</div>';

  // Birthdays
  document.getElementById('ov-bdays').innerHTML=birthdays.length?birthdays.slice(0,5).map(g=>{const d=bdayDays(g.date_of_birth);const color=d<=3?'#A32D2D':d<=7?'#BA7517':'#1D9E75';return`<div class="row"><div style="flex:1;min-width:0"><div class="rn">${escapeHtml(fn(g.full_name))}</div><div style="font-size:11px;color:#5F5E5A">${escapeHtml(g.email||'')}</div></div><div class="rd" style="color:${color};font-weight:500">in ${d}d</div></div>`;}).join(''):'<div class="empty">No birthdays in next 30 days.</div>';

  // In-house
  document.getElementById('ov-inhouse').innerHTML=inhouse.length?inhouse.map(s=>{const g=s.guests;const[bg,fg]=avc(g?.full_name||'');const stCls=s.status==='stayover'?'bm':s.status==='checked_in'?'be':'bg';return`<div class="row"><div class="av" style="background:${bg};color:${fg}">${ini(g?.full_name||'?')}</div><div style="flex:1;min-width:0"><div class="rn">${escapeHtml(g?.full_name||'')}</div><div style="font-size:11px;color:#5F5E5A">${escapeHtml(g?.nationality||'')} · ${escapeHtml(s.room_number||'')} · until ${fmtDate(s.departure_date)}</div></div><span class="badge ${stCls}">${s.status.replace('_',' ')}</span></div>`;}).join(''):'<div class="empty">No in-house guests.</div>';
}

// ============================================================================
// DASHBOARD
// ============================================================================
async function renderDashboard(){
  const el=document.getElementById('pane-dashboard');
  el.innerHTML='<div class="loading">Loading dashboard</div>';

  const guestIds = await getGuestIdsForProperty();
  const propFilter = STATE.selectedPropertyId !== 'all';

  let staySrcQ = sb.from('stays').select('source,channel_type');
  let nightsQ = sb.from('stays').select('nights');
  let arrQ = sb.from('stays').select('arrival_date').gte('arrival_date',new Date(Date.now()-30*864e5).toISOString().slice(0,10));
  if(propFilter){
    staySrcQ = staySrcQ.eq('property_id', STATE.selectedPropertyId);
    nightsQ = nightsQ.eq('property_id', STATE.selectedPropertyId);
    arrQ = arrQ.eq('property_id', STATE.selectedPropertyId);
  }

  let natQ = sb.from('guests').select('nationality');
  let typeQ = sb.from('guests').select('guest_type');
  if(guestIds){
    if(guestIds.length === 0){
      natQ = Promise.resolve({data: []});
      typeQ = Promise.resolve({data: []});
    } else {
      natQ = natQ.in('id', guestIds);
      typeQ = typeQ.in('id', guestIds);
    }
  }

  // Calculate property-specific stats if filtering
  let statsPromise, qualityPromise;
  if(propFilter && guestIds && guestIds.length > 0){
    statsPromise = (async () => {
      const stays = (await sb.from('stays').select('*').eq('property_id', STATE.selectedPropertyId)).data || [];
      const inHouse = stays.filter(s => ['checked_in','stayover'].includes(s.status)).length;
      const direct = stays.filter(s => s.channel_type === 'direct').length;
      const totalRev = stays.reduce((a,s) => a + (parseFloat(s.net_revenue_usd) || 0), 0);
      return {data:{total_guests:guestIds.length,in_house:inHouse,direct_bookings:direct,total_stays:stays.length,total_net_revenue:totalRev,repeat_guests:0}};
    })();
    qualityPromise = (async () => {
      const g = (await sb.from('guests').select('email,phone,date_of_birth,passport_number,national_id').in('id', guestIds)).data || [];
      const total = g.length || 1;
      const has_email = g.filter(x => x.email).length;
      const has_phone = g.filter(x => x.phone).length;
      const has_dob = g.filter(x => x.date_of_birth).length;
      const has_id = g.filter(x => x.passport_number || x.national_id).length;
      return {data:{total_guests:g.length,has_email,pct_email:Math.round(has_email/total*1000)/10,has_phone,pct_phone:Math.round(has_phone/total*1000)/10,has_dob,pct_dob:Math.round(has_dob/total*1000)/10,has_id,pct_id:Math.round(has_id/total*1000)/10}};
    })();
  } else if(propFilter && guestIds && guestIds.length === 0){
    statsPromise = Promise.resolve({data:{total_guests:0,in_house:0,direct_bookings:0,total_stays:0,total_net_revenue:0,repeat_guests:0}});
    qualityPromise = Promise.resolve({data:{total_guests:0,has_email:0,pct_email:0,has_phone:0,pct_phone:0,has_dob:0,pct_dob:0,has_id:0,pct_id:0}});
  } else {
    statsPromise = sb.from('v_dashboard_stats').select('*').single();
    qualityPromise = sb.from('v_data_quality').select('*').single();
  }

  const[stats,qualityRes,sourceRes,natRes,nightsRes,typeRes,arrRes]=await Promise.all([
    statsPromise,
    qualityPromise,
    staySrcQ,
    natQ,
    nightsQ,
    typeQ,
    arrQ
  ]);

  const s=stats.data||{};const q=qualityRes.data||{};

  const sourceCounts={};(sourceRes.data||[]).forEach(r=>{const k=r.source||'Unknown';sourceCounts[k]=(sourceCounts[k]||0)+1;});
  const natCounts={};(natRes.data||[]).forEach(r=>{const k=r.nationality||'Unknown';natCounts[k]=(natCounts[k]||0)+1;});
  const nightsCounts={};(nightsRes.data||[]).forEach(r=>{nightsCounts[r.nights]=(nightsCounts[r.nights]||0)+1;});

  // Build Property Performance card (always shown, even when filtered)
  const propPerf=propPerfRes.data||[];
  const accessibleIds=new Set(STATE.accessibleProperties.map(p=>p.id));
  const visiblePerf=propPerf.filter(p=>accessibleIds.has(p.property_id));

  const perfCardHtml=`<div class="card">
    <div class="ch"><div class="ct">Property Performance</div><div class="cs">side-by-side comparison · all time</div></div>
    <div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:13px;min-width:600px">
      <thead><tr style="border-bottom:1px solid var(--line-peach)">
        <th style="text-align:left;padding:10px 12px;font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--gray-text);font-weight:600">Property</th>
        <th style="text-align:right;padding:10px 12px;font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--gray-text);font-weight:600">Guests</th>
        <th style="text-align:right;padding:10px 12px;font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--gray-text);font-weight:600">Stays</th>
        <th style="text-align:right;padding:10px 12px;font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--gray-text);font-weight:600">Avg nights</th>
        <th style="text-align:right;padding:10px 12px;font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--gray-text);font-weight:600">In-house</th>
        <th style="text-align:right;padding:10px 12px;font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--gray-text);font-weight:600">Direct</th>
        <th style="text-align:right;padding:10px 12px;font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--gray-text);font-weight:600">Email %</th>
      </tr></thead>
      <tbody>${visiblePerf.map(p=>{
        const emailPct=p.unique_guests>0?Math.round(100*p.guests_with_email/p.unique_guests):0;
        const isCurrent=STATE.selectedPropertyId===p.property_id;
        return `<tr style="border-bottom:1px solid var(--line-peach);${isCurrent?'background:var(--kaani-cream-deep)':''}">
          <td style="padding:14px 12px;font-weight:600;color:var(--black)"><div style="display:flex;align-items:center;gap:8px"><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:var(--kaani-orange)"></span>${escapeHtml(p.property_name)}</div></td>
          <td style="padding:14px 12px;text-align:right;font-weight:600">${(p.unique_guests||0).toLocaleString()}</td>
          <td style="padding:14px 12px;text-align:right">${(p.total_stays||0).toLocaleString()}</td>
          <td style="padding:14px 12px;text-align:right">${p.avg_nights||'—'}</td>
          <td style="padding:14px 12px;text-align:right">${p.in_house||0}</td>
          <td style="padding:14px 12px;text-align:right">${p.direct_guests||0}</td>
          <td style="padding:14px 12px;text-align:right;color:${emailPct>=80?'var(--success)':emailPct>=50?'var(--warning)':'var(--danger)'};font-weight:600">${emailPct}%</td>
        </tr>`;
      }).join('')||'<tr><td colspan="7" style="padding:20px;text-align:center;color:var(--gray-text)">No property data yet — import guest lists to see comparison</td></tr>'}</tbody>
    </table></div>
  </div>`;

  const scopeLabel=STATE.selectedPropertyId
    ? STATE.properties.find(p=>p.id===STATE.selectedPropertyId)?.name
    : 'All Properties';

  el.innerHTML=`
    ${perfCardHtml}
    <div style="margin:24px 0 16px"><h3 style="font-family:'Inter Tight',sans-serif;font-size:18px;font-weight:700;color:var(--black);letter-spacing:-0.02em">${escapeHtml(scopeLabel)} — Detailed View</h3></div>
    <div class="sg">
      <div class="kpi"><div class="sl">Unique guests</div><div class="sv">${(s.total_guests||0).toLocaleString()}</div><div class="ss">${scopeLabel}</div><div class="kpi-bar"><div class="kpi-fill" style="width:100%;background:#1D9E75"></div></div></div>
      <div class="kpi"><div class="sl">Email capture</div><div class="sv">${q.pct_email||0}%</div><div class="ss">${q.has_email||0} of ${q.total_guests||0}</div><div class="kpi-bar"><div class="kpi-fill" style="width:${q.pct_email||0}%;background:#378ADD"></div></div></div>
      <div class="kpi"><div class="sl">Phone on file</div><div class="sv">${q.pct_phone||0}%</div><div class="ss">${q.has_phone||0} of ${q.total_guests||0}</div><div class="kpi-bar"><div class="kpi-fill" style="width:${q.pct_phone||0}%;background:#EF9F27"></div></div></div>
      <div class="kpi"><div class="sl">DOB on file</div><div class="sv">${q.pct_dob||0}%</div><div class="ss">${q.has_dob||0} of ${q.total_guests||0}</div><div class="kpi-bar"><div class="kpi-fill" style="width:${q.pct_dob||0}%;background:#1D9E75"></div></div></div>
    </div>
    <div class="g2">
      <div class="card"><div class="ch"><div class="ct">Booking source</div><div class="cs">all stays</div></div><div class="cw"><canvas id="srcChart"></canvas></div></div>
      <div class="card"><div class="ch"><div class="ct">Top nationalities</div><div class="cs">guest count</div></div><div class="cw"><canvas id="natChart"></canvas></div></div>
    </div>
    <div class="g2">
      <div class="card"><div class="ch"><div class="ct">Stay length distribution</div><div class="cs">stays by nights</div></div><div class="cw"><canvas id="nightsChart"></canvas></div></div>
      <div class="card"><div class="ch"><div class="ct">Tourist vs Local</div><div class="cs">guest type</div></div><div class="cw"><canvas id="typeChart"></canvas></div></div>
    </div>
    <div class="card"><div class="ch"><div class="ct">Daily arrivals</div><div class="cs">last 30 days</div></div><div class="cw" style="height:200px"><canvas id="arrChart"></canvas></div></div>
    <div class="card">
      <div class="ct" style="margin-bottom:12px">Data quality</div>
      <div style="font-size:13px;color:#5F5E5A;line-height:1.8">
        <div style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid #E8E6E0"><span>Email capture</span><strong style="color:#2C2C2A">${q.pct_email||0}%</strong></div>
        <div style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid #E8E6E0"><span>Phone capture</span><strong style="color:#2C2C2A">${q.pct_phone||0}%</strong></div>
        <div style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid #E8E6E0"><span>DOB capture</span><strong style="color:#2C2C2A">${q.pct_dob||0}%</strong></div>
        <div style="display:flex;justify-content:space-between;padding:6px 0"><span>ID/Passport on file</span><strong style="color:#2C2C2A">${q.pct_id||0}%</strong></div>
      </div>
    </div>
    <div class="card" id="insights-card">
      <div class="ct" style="margin-bottom:12px">Key insights for action</div>
      <div id="insights-body"></div>
    </div>`;

  // Render charts
  const srcOrder=Object.entries(sourceCounts).sort((a,b)=>b[1]-a[1]).slice(0,6);
  new Chart(document.getElementById('srcChart'),{type:'doughnut',data:{labels:srcOrder.map(([n])=>srcShort(n)),datasets:[{data:srcOrder.map(([,v])=>v),backgroundColor:['#1D9E75','#378ADD','#EF9F27','#5DCAA5','#D85A30','#AFA9EC'],borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,cutout:'65%',plugins:{legend:{position:'bottom',labels:{font:{size:11}}}}}});

  const natOrder=Object.entries(natCounts).sort((a,b)=>b[1]-a[1]).slice(0,10);
  new Chart(document.getElementById('natChart'),{type:'bar',data:{labels:natOrder.map(([n])=>n),datasets:[{data:natOrder.map(([,v])=>v),backgroundColor:'#1D9E75',borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{beginAtZero:true,ticks:{font:{size:11}}},y:{ticks:{font:{size:11}}}}}});

  const nKeys=Object.keys(nightsCounts).map(Number).sort((a,b)=>a-b);
  new Chart(document.getElementById('nightsChart'),{type:'bar',data:{labels:nKeys.map(n=>n+'n'),datasets:[{data:nKeys.map(n=>nightsCounts[n]),backgroundColor:'#9FE1CB',borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{font:{size:11}}},x:{ticks:{font:{size:11}}}}}});

  // Tourist vs Local donut
  const typeCounts={tourist:0,local:0};
  (typeRes.data||[]).forEach(g=>{if(g.guest_type==='tourist')typeCounts.tourist++;else if(g.guest_type==='local')typeCounts.local++;});
  const totalType=typeCounts.tourist+typeCounts.local;
  new Chart(document.getElementById('typeChart'),{type:'doughnut',data:{labels:['Tourist','Local'],datasets:[{data:[typeCounts.tourist,typeCounts.local],backgroundColor:['#1D9E75','#AFA9EC'],borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,cutout:'65%',plugins:{legend:{position:'bottom',labels:{font:{size:11},generateLabels:(c)=>{const d=c.data;return d.labels.map((l,i)=>({text:l+' '+d.datasets[0].data[i]+(totalType?' ('+Math.round(d.datasets[0].data[i]/totalType*100)+'%)':''),fillStyle:d.datasets[0].backgroundColor[i],strokeStyle:d.datasets[0].backgroundColor[i],index:i}));}}}}}});

  // Daily arrivals line chart
  const arrCounts={};(arrRes.data||[]).forEach(s=>{if(s.arrival_date)arrCounts[s.arrival_date]=(arrCounts[s.arrival_date]||0)+1;});
  const arrSorted=Object.entries(arrCounts).sort();
  const arrLabels=arrSorted.map(([d])=>{const dt=new Date(d);return String(dt.getDate()).padStart(2,'0')+' '+MONTHS[dt.getMonth()];});
  const arrVals=arrSorted.map(([,v])=>v);
  new Chart(document.getElementById('arrChart'),{type:'line',data:{labels:arrLabels,datasets:[{data:arrVals,borderColor:'#1D9E75',backgroundColor:'rgba(29,158,117,0.12)',fill:true,tension:0.3,pointBackgroundColor:'#1D9E75',pointRadius:4,pointHoverRadius:6,borderWidth:2.5}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{font:{size:11}}},x:{ticks:{font:{size:10},maxRotation:0}}}}});

  // Generate insights dynamically
  generateInsights(s,q,sourceCounts,natCounts,typeCounts,nightsCounts);
}

function generateInsights(stats,quality,sources,nats,types,nightsDist){
  const insights=[];
  const totalStays=Object.values(sources).reduce((a,b)=>a+b,0);
  const totalGuests=Object.values(nats).reduce((a,b)=>a+b,0);
  if(totalStays>0){
    const sortedSources=Object.entries(sources).sort((a,b)=>b[1]-a[1]);
    const topSrc=sortedSources[0];
    const topPct=Math.round(topSrc[1]/totalStays*100);
    if(topPct>40){
      insights.push({title:'Source concentration risk.',body:topPct+'% of stays come through one channel ('+srcShort(topSrc[0])+'). Diversifying source channels reduces risk.'});
    }
  }
  if(totalGuests>0){
    const sortedNats=Object.entries(nats).sort((a,b)=>b[1]-a[1]);
    const topNat=sortedNats[0];
    const topNatPct=Math.round(topNat[1]/totalGuests*100);
    if(topNatPct>30){
      insights.push({title:topNat[0]+' market dominance.',body:topNatPct+'% of guests are from '+topNat[0]+'. Consider language-specific content and culturally tailored welcome touches.'});
    }
  }
  const totalNights=Object.values(nightsDist).reduce((a,b)=>a+b,0);
  const longStays=Object.entries(nightsDist).filter(([n])=>+n>=7).reduce((a,[,v])=>a+v,0);
  if(totalNights>0&&longStays/totalNights>0.3){
    insights.push({title:'Long-stay loyalty opportunity.',body:Math.round(longStays/totalNights*100)+'% of stays are 7+ nights — these are your most engaged guests and the prime targets for win-back campaigns.'});
  }
  if(quality.pct_email<90){
    insights.push({title:'Email gap.',body:'Only '+(quality.pct_email||0)+'% of guests have an email on file. Capturing missing emails at check-in is the single biggest lift to your marketing reach.'});
  }
  const directCount=Object.entries(sources).filter(([s])=>s&&(s.toUpperCase().includes('DIRECT'))).reduce((a,[,v])=>a+v,0);
  if(totalStays>0){
    const directPct=Math.round(directCount/totalStays*100);
    if(directPct<25){
      insights.push({title:'Direct booking upside.',body:'Only '+directPct+'% of bookings are direct vs 25–30% benchmark. The win-back campaign in your action plan targets this gap.'});
    }
  }
  document.getElementById('insights-body').innerHTML=insights.length?insights.map(i=>'<div style="padding:10px 0;border-bottom:1px solid #E8E6E0;font-size:13px;color:#5F5E5A;line-height:1.7"><strong style="color:#2C2C2A;font-weight:500">'+i.title+'</strong> '+i.body+'</div>').join(''):'<div style="font-size:13px;color:#5F5E5A">More insights will appear as your data grows.</div>';
}

// ============================================================================
// GUESTS
// ============================================================================
async function renderGuestsPane(){
  const el=document.getElementById('pane-guests');
  el.innerHTML=`
    <div class="sg" id="g-stats"></div>
    <div class="toolbar">
      <input type="text" id="gsearch" placeholder="Search name, email, nationality..." />
      <select id="gfst"><option value="">All statuses</option><option value="prospect">Prospect</option><option value="first_time">First-time</option><option value="repeat">Repeat</option><option value="vip">VIP</option><option value="lapsed">Lapsed</option></select>
      <select id="gfnat"><option value="">All nationalities</option></select>
    </div>
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

  // Populate nationality filter
  const{data:nats}=await sb.from('guests').select('nationality').not('nationality','is',null).limit(2000);
  const uniq=[...new Set((nats||[]).map(n=>n.nationality))].sort();
  const ne=document.getElementById('gfnat');
  uniq.forEach(n=>{const o=document.createElement('option');o.value=n;o.text=n;ne.appendChild(o);});

  document.getElementById('gsearch').oninput=()=>{STATE.guestsPage=0;loadGuests();};
  document.getElementById('gfst').onchange=()=>{STATE.guestsPage=0;loadGuests();};
  document.getElementById('gfnat').onchange=()=>{STATE.guestsPage=0;loadGuests();};

  loadGuests();
}

let guestSearchTimer;
async function loadGuests(){
  clearTimeout(guestSearchTimer);
  guestSearchTimer=setTimeout(async()=>{
    const q=document.getElementById('gsearch').value.trim();
    const fst=document.getElementById('gfst').value;
    const fnat=document.getElementById('gfnat').value;
    const from=STATE.guestsPage*STATE.guestsPerPage;
    const to=from+STATE.guestsPerPage-1;

    let query=sb.from('guests').select('*',{count:'exact'}).order('last_stay_date',{ascending:false,nullsFirst:false}).range(from,to);
    if(STATE.selectedPropertyId !== 'all'){
      const propGuestIds = await getGuestIdsForProperty();
      if(propGuestIds.length === 0){
        STATE.currentGuests = [];
        renderGuestList(0);
        return;
      }
      query = query.in('id', propGuestIds);
    }
    if(q){query=query.or(`full_name.ilike.%${q}%,email.ilike.%${q}%,nationality.ilike.%${q}%`);}
    if(fst)query=query.eq('lead_status',fst);
    if(fnat)query=query.eq('nationality',fnat);

    const{data,error,count}=await query;
    if(error){toast('Load error: '+error.message,true);return;}

    STATE.currentGuests=data||[];
    renderGuestList(count||0);
  },200);
}

function renderGuestList(totalCount){
  const el=document.getElementById('glist');
  if(!STATE.currentGuests.length){el.innerHTML='<div class="empty">No guests match your filters</div>';return;}
  el.innerHTML=STATE.currentGuests.map(g=>{
    const[bg,fg]=avc(g.full_name);
    const age=calcAge(g.date_of_birth);
    return`<div class="gcard" onclick="openGuestPanel('${g.id}')">
      <div style="display:flex;align-items:center;gap:11px">
        <div class="av" style="background:${bg};color:${fg}">${ini(g.full_name)}</div>
        <div style="flex:1;min-width:0">
          <div style="font-size:14px;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${escapeHtml(g.full_name)}</div>
          <div style="font-size:12px;color:#5F5E5A">${g.date_of_birth?'DOB: '+fmtDate(g.date_of_birth)+(age?' · age '+age:''):escapeHtml(g.email||'no email')}</div>
        </div>
        <div style="display:flex;flex-direction:column;align-items:flex-end;gap:4px;flex-shrink:0">
          <span class="badge ${g.lead_status==='vip'?'bu':g.lead_status==='repeat'?'be':g.lead_status==='lapsed'?'bg':'bm'}">${g.lead_status.replace('_',' ')}</span>
          <span style="font-size:11px;color:#5F5E5A">${g.total_stays} ${g.total_stays===1?'stay':'stays'}</span>
        </div>
      </div>
      <div style="display:flex;gap:10px;margin-top:9px;padding-top:9px;border-top:1px solid #E8E6E0;flex-wrap:wrap;font-size:12px">
        <span style="color:#5F5E5A"><strong style="color:#2C2C2A">${escapeHtml(g.nationality||'—')}</strong></span>
        <span style="color:#5F5E5A">Last stay <strong style="color:#2C2C2A">${g.last_stay_date?fmtDate(g.last_stay_date):'—'}</strong></span>
        ${g.guest_type==='local'?'<span class="badge bl">Local</span>':''}
      </div>
    </div>`;
  }).join('');
  renderPagination(totalCount);
}

function renderPagination(total){
  const totalPages=Math.ceil(total/STATE.guestsPerPage);
  if(totalPages<=1){document.getElementById('g-pagination').innerHTML='';return;}
  const p=STATE.guestsPage;
  let html=`<button onclick="changeGuestPage(${p-1})" ${p===0?'disabled':''}>‹ Prev</button>`;
  html+=`<button class="on">${p+1}</button>`;
  html+=`<button onclick="changeGuestPage(${p+1})" ${p>=totalPages-1?'disabled':''}>Next ›</button>`;
  html+=`<span style="font-size:11px;color:#5F5E5A;align-self:center;margin-left:8px">Page ${p+1} of ${totalPages} · ${total.toLocaleString()} guests</span>`;
  document.getElementById('g-pagination').innerHTML=html;
}
function changeGuestPage(p){STATE.guestsPage=p;loadGuests();}

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
        <div class="dr"><span class="dk">Email</span><span class="dv">${escapeHtml(g.email||'—')}</span></div>
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
async function renderActionsPane(){
  const el=document.getElementById('pane-actions');
  el.innerHTML=`
    <div class="nav" style="margin-bottom:14px" id="action-nav">
      <button class="nb on" onclick="renderAction('review_request',this)">Review request</button>
      <button class="nb" onclick="renderAction('winback',this)">Win-back</button>
      <button class="nb" onclick="renderAction('upsell',this)">Upsell</button>
      <button class="nb" onclick="renderAction('birthday',this)">Birthdays</button>
      <button class="nb" onclick="renderAction('seasonal',this)">Seasonal</button>
      <button class="nb" onclick="renderChecklist(this)">Checklist</button>
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

  let guests=[];
  let segmentLabel='';

  const actionPropGuestIds = await getGuestIdsForProperty();
  const propLabel = STATE.selectedPropertyId === 'all' ? '' : ' (' + (STATE.properties.find(p=>p.id==STATE.selectedPropertyId)?.name||'') + ')';
  function filterByProp(guestList){
    if(!actionPropGuestIds) return guestList;
    const idSet = new Set(actionPropGuestIds);
    return guestList.filter(g => idSet.has(g.id));
  }

  if(key==='review_request'){
    let q = sb.from('stays').select('*,guests(*)').eq('status','checked_out').gte('departure_date',new Date(Date.now()-3*864e5).toISOString().slice(0,10)).order('departure_date',{ascending:false});
    if(STATE.selectedPropertyId !== 'all') q = q.eq('property_id', STATE.selectedPropertyId);
    const{data}=await q;
    guests=(data||[]).filter(s=>s.guests?.email&&s.guests?.marketing_consent).map(s=>s.guests);
    segmentLabel='Recent checkouts (last 3 days)' + propLabel;
  }else if(key==='winback'){
    let q = sb.from('guests').select('*').not('email','is',null).eq('marketing_consent',true).in('lead_status',['first_time','lapsed']).order('last_stay_date',{ascending:false}).limit(200);
    if(actionPropGuestIds){
      if(actionPropGuestIds.length === 0){ guests = []; segmentLabel = 'No guests' + propLabel; }
      else { q = q.in('id', actionPropGuestIds); const{data}=await q; guests=data||[]; segmentLabel='First-time + lapsed guests with emails' + propLabel; }
    } else {
      const{data}=await q; guests=data||[]; segmentLabel='First-time + lapsed guests with emails';
    }
  }else if(key==='upsell'){
    let q = sb.from('stays').select('*,guests(*)').in('status',['checked_in','stayover']).gte('nights',3);
    if(STATE.selectedPropertyId !== 'all') q = q.eq('property_id', STATE.selectedPropertyId);
    const{data}=await q;
    guests=(data||[]).filter(s=>s.guests?.email&&s.guests?.marketing_consent).map(s=>s.guests);
    segmentLabel='Current stayovers (3+ nights with email)' + propLabel;
  }else if(key==='birthday'){
    let q = sb.from('guests').select('*').not('email','is',null).not('date_of_birth','is',null).eq('marketing_consent',true).limit(500);
    if(actionPropGuestIds){
      if(actionPropGuestIds.length === 0){ guests = []; segmentLabel = 'No guests' + propLabel; }
      else { q = q.in('id', actionPropGuestIds); const{data}=await q; guests=(data||[]).filter(g=>{const d=bdayDays(g.date_of_birth);return d!==null&&d<=30;}).sort((a,b)=>bdayDays(a.date_of_birth)-bdayDays(b.date_of_birth)); segmentLabel='Birthdays in next 30 days' + propLabel; }
    } else {
      const{data}=await q;
      guests=(data||[]).filter(g=>{const d=bdayDays(g.date_of_birth);return d!==null&&d<=30;}).sort((a,b)=>bdayDays(a.date_of_birth)-bdayDays(b.date_of_birth));
      segmentLabel='Birthdays in next 30 days';
    }
  }else if(key==='seasonal'){
    let q = sb.from('guests').select('*').not('email','is',null).eq('marketing_consent',true).limit(2000);
    if(actionPropGuestIds){
      if(actionPropGuestIds.length === 0){ guests = []; segmentLabel = 'No guests' + propLabel; }
      else { q = q.in('id', actionPropGuestIds); const{data}=await q; guests=data||[]; segmentLabel='All guests with marketing consent' + propLabel; }
    } else {
      const{data}=await q; guests=data||[]; segmentLabel='All guests with marketing consent';
    }
  }

  // Deduplicate by id
  const seen=new Set();const unique=[];
  guests.forEach(g=>{if(!seen.has(g.id)){seen.add(g.id);unique.push(g);}});
  guests=unique;

  document.getElementById('action-content').innerHTML=`
    <div class="card">
      <div class="ch"><div class="ct">${escapeHtml(tpl.name)}</div><span class="badge be">${escapeHtml(tpl.description||'')}</span></div>
      <div class="sec">Segment: ${escapeHtml(segmentLabel)} (${guests.length} guests)</div>
      <div style="margin-bottom:14px;max-height:200px;overflow-y:auto">
        ${guests.slice(0,5).map(g=>`<div class="row"><div style="flex:1;min-width:0;color:#5F5E5A;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${escapeHtml(g.full_name)}</div><div style="font-size:12px;color:#2C2C2A">${escapeHtml(g.email||'')}</div></div>`).join('')+(guests.length>5?`<div style="font-size:12px;color:#5F5E5A;padding-top:8px">+${guests.length-5} more</div>`:'')}
      </div>
      <div class="sec">Subject</div>
      <div class="tpl">${escapeHtml(tpl.subject||'')}</div>
      <div class="sec">Body template</div>
      <div class="tpl" id="tpl-body">${escapeHtml(tpl.body||'')}</div>
      <div class="btns">
        <button class="btn btnp" onclick="copyToClipboard(document.getElementById('tpl-body').innerText,'Template copied')">Copy template</button>
        <button class="btn" onclick='copyToClipboard(${JSON.stringify(guests.map(g=>g.email).join(", "))},"Emails copied")'>Copy ${guests.length} emails</button>
        <button class="btn" onclick="logCampaignSend('${key}',${guests.length})">Log as sent</button>
      </div>
      <div style="margin-top:10px;font-size:12px;color:#5F5E5A">Templates use {{first_name}} and {{property_name}} — replace before sending. Edit the template in the Templates tab.</div>
    </div>`;
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

  const propGuestIds = await getGuestIdsForProperty();
  let allQ = sb.from('guests').select('id,full_name,email,nationality,lead_status,date_of_birth,marketing_consent').not('email','is',null).eq('marketing_consent',true).limit(5000);
  if(propGuestIds){
    if(propGuestIds.length === 0){
      el.innerHTML='<div class="empty">No guests yet for this property</div>';
      return;
    }
    allQ = allQ.in('id', propGuestIds);
  }
  const{data:all}=await allQ;
  const list=all||[];

  // Get direct booker IDs separately
  const{data:directStays}=await sb.from('stays').select('guest_id').eq('channel_type','direct');
  const directIds=new Set((directStays||[]).map(s=>s.guest_id));

  const direct=list.filter(g=>directIds.has(g.id));
  const repeat=list.filter(g=>['repeat','vip'].includes(g.lead_status));
  const withDOB=list.filter(g=>g.date_of_birth);

  el.innerHTML=`
    <div class="sg">
      <div class="sm"><div class="sl">Total contactable</div><div class="sv">${list.length}</div></div>
      <div class="sm"><div class="sl">Direct bookers</div><div class="sv">${direct.length}</div></div>
      <div class="sm"><div class="sl">Repeat & VIP</div><div class="sv">${repeat.length}</div></div>
      <div class="sm"><div class="sl">With DOB</div><div class="sv">${withDOB.length}</div></div>
    </div>
    <div style="margin-bottom:12px"><button class="btn btnp" onclick='copyToClipboard(${JSON.stringify(list.map(g=>g.email).join(", "))},"All emails copied")'>Copy all ${list.length} emails</button></div>
    <div id="mkt-groups"></div>`;

  const groups=[
    {label:'All guests with email',detail:`${list.length} contacts (consented)`,emails:list.map(g=>g.email)},
    {label:'Direct bookings',detail:`${direct.length} contacts — best for loyalty offers`,emails:direct.map(g=>g.email)},
    {label:'Repeat & VIP guests',detail:`${repeat.length} contacts — your most valuable`,emails:repeat.map(g=>g.email)},
    {label:'Guests with DOB on file',detail:`${withDOB.length} contacts — birthday campaign targets`,emails:withDOB.map(g=>g.email)},
  ];
  document.getElementById('mkt-groups').innerHTML=groups.map((g,i)=>`
    <div class="card" style="padding:13px 15px;display:flex;align-items:center;gap:12px;flex-wrap:wrap">
      <div style="flex:1;min-width:0"><div style="font-size:13px;font-weight:500">${escapeHtml(g.label)}</div><div style="font-size:11px;color:#5F5E5A;margin-top:3px">${escapeHtml(g.detail)}</div></div>
      <button class="btn" onclick='copyToClipboard(${JSON.stringify(g.emails.join(", "))},"Emails copied")'>Copy emails</button>
    </div>`).join('');
}

// ============================================================================
// TEMPLATES — editable, no code change needed
// ============================================================================
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
  const isAll=propIds.length===STATE.properties.length;
  let q=sb.from('guests').select('*').order('last_stay_date',{ascending:false}).limit(50000);
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
  let q=sb.from('stays').select('*,guests(full_name,email,nationality),properties(name)').in('property_id',propIds).limit(50000);
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
  const{data}=await sb.from('guests').select('*').gte('total_stays',2).order('total_stays',{ascending:false}).limit(50000);
  const rows=(data||[]).map(g=>[g.full_name,g.email,g.nationality,g.total_stays,g.total_nights,g.total_revenue_usd,g.first_stay_date,g.last_stay_date]);
  downloadCSV('kaani_repeat_guests_'+new Date().toISOString().slice(0,10)+'.csv',['Name','Email','Nationality','Stays','Nights','Revenue','First Stay','Last Stay'],rows);
}

async function downloadCampaignReport(){
  const{data}=await sb.from('campaigns').select('*').order('sent_at',{ascending:false}).limit(1000);
  const rows=(data||[]).map(c=>[c.name,c.campaign_type,c.status,c.sent_at,c.emails_sent,c.responses_received,c.bookings_attributed,c.revenue_attributed]);
  downloadCSV('kaani_campaigns_'+new Date().toISOString().slice(0,10)+'.csv',['Name','Type','Status','Sent At','Emails Sent','Responses','Bookings','Revenue'],rows);
}

// ============================================================================
// IMPORT — Ezee CSV with proper deduplication
// ============================================================================
function renderImportPane(){
  document.getElementById('pane-import').innerHTML=`
    <div class="card" style="max-width:560px">
      <div class="ct" style="margin-bottom:8px">Import Ezee guest list</div>
      <div class="cd">Upload an Ezee daily guest list CSV. The system matches by passport / national ID — guests are never duplicated. New stays are added; existing guest profiles get updated.</div>
      <div style="margin-bottom:12px">
        <label style="font-size:13px;color:#5F5E5A;display:block;margin-bottom:6px">Property</label>
        <select id="imp-prop" style="width:100%;padding:9px 12px;border-radius:8px;border:1px solid #E8E6E0;font-family:inherit">
          ${STATE.properties.map(p=>`<option value="${p.id}">${escapeHtml(p.name)}</option>`).join('')}
        </select>
      </div>
      <div class="izone"><div style="font-size:24px;margin-bottom:8px">+</div>Drop CSV or <label>browse to upload<input type="file" accept=".csv" style="display:none" onchange="processImport(event)"></label></div>
      <div id="imp-msg" style="margin-top:14px;font-size:13px"></div>
      <div id="imp-progress" style="margin-top:8px"></div>
    </div>`;
}

function parseCSV(text){
  // Strip BOM (Byte Order Mark) if present - common in Windows CSV exports
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
  // Strip BOM from individual headers and trim, also remove any quotes
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

function parseEzeeDate(d){if(!d)return null;const m=d.match(/(\d{1,2})-(\d{1,2})-(\d{2,4})/);if(!m)return null;const[,dd,mm,yy]=m;const y=yy.length===2?'20'+yy:yy;return`${y}-${mm.padStart(2,'0')}-${dd.padStart(2,'0')}`;}

function classifyChannel(source){
  if(!source)return'other';
  const s=source.toUpperCase();
  if(s==='DIRECT BOOKING'||s.includes('DIRECT'))return'direct';
  if(s.includes('BOOKING.COM')||s.includes('AGODA')||s.includes('EXPEDIA')||s.includes('MAKE MY TRIP')||s.includes('HOTELS.COM'))return'ota';
  if(s.includes('DESTINOS')||s.includes('TOUR'))return'tour_operator';
  if(s.includes('WALK'))return'walk_in';
  return'other';
}

async function processImport(e){
  const file=e.target.files[0];if(!file)return;
  const propId=parseInt(document.getElementById('imp-prop').value);
  const msg=document.getElementById('imp-msg');
  const prog=document.getElementById('imp-progress');
  msg.textContent='Reading file...';msg.style.color='#5F5E5A';

  const text=await file.text();
  const{headers,rows}=parseCSV(text);
  const idx={};headers.forEach((h,i)=>idx[h]=i);

  const required=['Identity','Guest Name','Arrival','Departure','Nights'];
  for(const r of required){if(!(r in idx)){msg.textContent='Missing column: '+r+'. Found columns: '+headers.join(', ');msg.style.color='#A32D2D';return;}}

  msg.textContent=`Processing ${rows.length} rows...`;
  let guestsNew=0,guestsUpdated=0,staysNew=0,errors=0;

  for(let i=0;i<rows.length;i++){
    const r=rows[i];
    const identity=r[idx.Identity]?.trim();if(!identity)continue;

    const isNationalId=identity.toLowerCase().includes('national id');
    const passport=isNationalId?null:identity.replace(/^Passport-/i,'');
    const nationalId=isNationalId?identity.replace(/^National ID-/i,''):null;

    const name=r[idx['Guest Name']]?.trim()||'';
    let email=r[idx.Email]?.trim()||null;
    if(email&&isPlaceholderEmail(email))email=null;  // Skip placeholder/operator emails
    const nationality=r[idx.Nationality]?.trim()||null;
    const address=r[idx.Address]?.trim()||null;
    const dob=extractDOB(address);

    const nameParts=name.replace(/^(Mr\.|Ms\.|Mrs\.|Dr\.)\s+/i,'').trim().split(/\s+/);
    const firstName=nameParts[0]||null;
    const lastName=nameParts.length>1?nameParts.slice(-1)[0]:null;

    // Find or create guest
    let guestId;
    let existing=null;
    if(passport){
      const{data}=await sb.from('guests').select('id').eq('passport_number',passport).maybeSingle();
      existing=data;
    }
    if(!existing&&nationalId){
      const{data}=await sb.from('guests').select('id').eq('national_id',nationalId).maybeSingle();
      existing=data;
    }
    if(!existing&&email&&!isPlaceholderEmail(email)){
      const{data}=await sb.from('guests').select('id').ilike('email',email).maybeSingle();
      existing=data;
    }

    if(existing){
      guestId=existing.id;
      // Update with any new info
      const updates={full_name:name,first_name:firstName,last_name:lastName};
      if(email)updates.email=email;
      if(nationality)updates.nationality=nationality;
      if(dob)updates.date_of_birth=dob;
      if(passport)updates.passport_number=passport;
      if(nationalId)updates.national_id=nationalId;
      const{error}=await sb.from('guests').update(updates).eq('id',guestId);
      if(error){errors++;continue;}
      guestsUpdated++;
    }else{
      const guestType=r[idx['Guest Type']]?.includes('Local')?'local':'tourist';
      const{data,error}=await sb.from('guests').insert({full_name:name,first_name:firstName,last_name:lastName,email,nationality,date_of_birth:dob,passport_number:passport,national_id:nationalId,guest_type:guestType,created_by:STATE.user.id}).select('id').single();
      if(error){errors++;continue;}
      guestId=data.id;
      guestsNew++;
    }

    // Add stay (deduplicate by guest+arrival+property)
    const arrival=parseEzeeDate(r[idx.Arrival]);
    const departure=parseEzeeDate(r[idx.Departure]);
    const nights=parseInt(r[idx.Nights])||0;

    const{data:existingStay}=await sb.from('stays').select('id').eq('guest_id',guestId).eq('property_id',propId).eq('arrival_date',arrival).maybeSingle();

    const source=r[idx.Source]?.trim()||null;
    const status=r[idx.Status]?.toLowerCase().replace(' ','_').replace('checked_in','checked_in').replace('checked_out','checked_out')||'reservation';
    const paxStr=r[idx.Pax]||'';
    const paxMatch=paxStr.match(/(\d+)\s*\(A\)\s*\/\s*(\d+)\s*\(C\)/);
    const paxA=paxMatch?+paxMatch[1]:1;const paxC=paxMatch?+paxMatch[2]:0;

    const stayData={guest_id:guestId,property_id:propId,arrival_date:arrival,departure_date:departure,nights,room_number:r[idx.Room]?.trim()||null,rate_type:r[idx['Rate Type']]?.trim()||null,pax_adults:paxA,pax_children:paxC,source,channel_type:classifyChannel(source),status,reservation_remark:r[idx['Reservation Remark']]?.trim()||null};

    if(existingStay){
      await sb.from('stays').update(stayData).eq('id',existingStay.id);
    }else{
      const{error}=await sb.from('stays').insert(stayData);
      if(error){errors++;continue;}
      staysNew++;
    }

    if(i%10===0){prog.textContent=`Processed ${i+1} of ${rows.length}...`;}
  }

  msg.textContent=`Done: ${guestsNew} new guests, ${guestsUpdated} updated, ${staysNew} new stays, ${errors} errors.`;
  msg.style.color='#0F6E56';
  prog.textContent='';
  await loadInitialData();
  toast('Import complete');
}

// ============================================================================
// ADMIN
// ============================================================================
async function renderAdminPane(){
  if(STATE.profile.role!=='admin'){document.getElementById('pane-admin').innerHTML='<div class="empty">Admin access required</div>';return;}
  const{data:users}=await sb.from('user_profiles').select('*').order('created_at');
  const{data:dups}=await sb.from('v_potential_duplicates').select('*').limit(50);

  document.getElementById('pane-admin').innerHTML=`
    <div class="card">
      <div class="ct" style="margin-bottom:10px">Team members</div>
      <div class="cd">Add team members by sharing the login URL — they sign up themselves with email/password, then you assign their role here.</div>
      ${(users||[]).map(u=>`<div class="row">
        <div style="flex:1;min-width:0"><div class="rn">${escapeHtml(u.full_name||u.user_id)}</div><div style="font-size:11px;color:#5F5E5A">${u.user_id===STATE.user.id?'(you)':''}</div></div>
        <select onchange="updateUserRole('${u.user_id}',this.value)" ${u.user_id===STATE.user.id?'disabled':''} style="font-size:12px;padding:6px 10px;border-radius:8px;border:1px solid var(--border-subtle);font-family:inherit;background:#fff">
          <option value="staff" ${u.role==='staff'?'selected':''}>Staff</option>
          <option value="manager" ${u.role==='manager'?'selected':''}>Manager</option>
          <option value="admin" ${u.role==='admin'?'selected':''}>Admin</option>
        </select>
        <button class="btn" style="font-size:11px;padding:6px 12px" onclick="manageProperties('${u.user_id}','${escapeHtml(u.full_name||'')}')">Properties</button>
      </div>`).join('')}
    </div>
    <div class="card">
      <div class="ct" style="margin-bottom:10px">Potential duplicate guests (${dups?.length||0})</div>
      <div class="cd">These guests have similar details — review and merge if they're the same person.</div>
      ${(dups||[]).slice(0,20).map(d=>`<div class="row"><div style="flex:1;min-width:0"><div class="rn">${escapeHtml(d.name_1)} ↔ ${escapeHtml(d.name_2)}</div><div style="font-size:11px;color:#5F5E5A">${escapeHtml(d.match_reason)}</div></div></div>`).join('')||'<div class="empty">No duplicates detected</div>'}
    </div>
    <div class="card">
      <div class="ct" style="margin-bottom:10px">Quick actions</div>
      <div class="btns">
        <button class="btn" onclick="downloadGuestReport()">Backup all guests (CSV)</button>
        <button class="btn" onclick="downloadStaysReport()">Backup all stays (CSV)</button>
      </div>
    </div>`;
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
  sp.innerHTML=`<div class="panel-wrap" onclick="closePanelOverlay(event)">
    <div class="panel" onclick="event.stopPropagation()">
      <button class="pc" onclick="closeAdminPanel()">✕</button>
      <div class="pn">Property access</div>
      <div class="ps">${escapeHtml(userName||'team member')}</div>
      <div class="cd">Select which properties this user can access. Leave empty to grant access to all properties.</div>
      <div style="margin-top:14px">
        ${STATE.properties.map(p=>`<label style="display:flex;align-items:center;gap:10px;padding:11px 14px;border:1px solid var(--line-peach);border-radius:10px;margin-bottom:8px;cursor:pointer;background:var(--kaani-cream)">
          <input type="checkbox" id="prop-chk-${p.id}" ${current.includes(p.id)?'checked':''} style="width:16px;height:16px;accent-color:var(--kaani-orange);cursor:pointer">
          <span style="font-weight:500">${escapeHtml(p.name)}</span>
          <span style="margin-left:auto;font-size:11px;color:var(--gray-text)">${escapeHtml(p.code)}</span>
        </label>`).join('')}
      </div>
      <button class="btn btnp" style="width:100%;margin-top:14px" onclick="savePropertyAccess('${userId}')">Save access</button>
    </div>
  </div>`;
}

function closeAdminPanel(){document.getElementById('sp').style.display='none';}
function closePanelOverlay(e){if(e&&!e.target.classList.contains('panel-wrap'))return;closeAdminPanel();}

async function savePropertyAccess(userId){
  const selected=[];
  STATE.properties.forEach(p=>{
    const cb=document.getElementById('prop-chk-'+p.id);
    if(cb&&cb.checked)selected.push(p.id);
  });
  const{error}=await sb.from('user_profiles').update({property_ids:selected.length>0?selected:null}).eq('user_id',userId);
  if(error){toast(error.message,true);return;}
  toast('Property access updated');
  closeAdminPanel();
  renderAdminPane();
}

// ------------- INIT -------------
checkSession();
document.getElementById('login-password').addEventListener('keypress',e=>{if(e.key==='Enter')doLogin();});
