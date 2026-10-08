export const PORTAL_GROUPS = [
  {label:'Workspace',keys:['home']},
  {label:'Boards',keys:['board','wip','design-board']},
  {label:'Team',keys:['attendance','holidays','mileage','installer']},
  {label:'Tools',keys:['order-panels','van-estimator','rams','social-post','description-pull','pro-forma','igloo']},
  {label:'System',keys:['corebridge-explorer','notifications']}
];
export const MODULE_META = {
  board:['Installation Board','Plan installations, see availability and keep each job moving.','Boards'],
  wip:['WIP','Drag jobs between production days and status lanes in one shared board.','Boards'],
  'design-board':['Design Board','Allocate artwork, track approvals and keep customer responses in view.','Boards'],
  attendance:['Attendance','Review the month, then select a person to inspect or adjust their clockings.','Team'],
  holidays:['Holidays','Plan team availability, review requests and manage leave allowances.','Team'],
  mileage:['Mileage','Review journeys, monthly claims and outstanding submissions.','Team'],
  installer:['Subcontractors','Find the right installer, coverage and contact details for each job.','Team'],
  'order-panels':['Order Panels','Build your materials list and review sheet layouts before ordering.','Tools'],
  'van-estimator':['Vehicle Pricing','Choose a vehicle, plan its graphics and build a clear proposal.','Tools'],
  rams:['RAMS','Prepare risk assessments and method statements from your job details.','Tools'],
  'rams-logic':['RAMS Library','Maintain the risk, method and rescue options available to the team.','Tools'],
  'social-post':['Social Post','Turn completed work into a post, using your photos and chosen voice.','Tools'],
  'description-pull':['Description Pull','Find the source order and copy the customer description you need.','Tools'],
  'pro-forma':['Pro-Forma','Pull an order, check the details, choose a deposit and preview your invoice.','Tools'],
  'pro-forma-template':['Invoice Template','Maintain the approved invoice layout.','Tools'],
  'corebridge-explorer':['CoreBridge APIs','Look up source records and inspect the response.','System'],
  notifications:['Notifications','Follow the updates that need your attention.','System'],
  igloo:['IGLOO','Browse and maintain your product catalogue.','Tools']
};
export function portalWeekDays(today, offset=0) {
  const start=new Date(today+'T12:00:00Z');
  start.setUTCDate(start.getUTCDate()-((start.getUTCDay()+6)%7)+offset*7);
  return Array.from({length:5},(_,i)=>{
    const d=new Date(start);d.setUTCDate(d.getUTCDate()+i);
    return {id:d.toISOString().slice(0,10),label:d.toLocaleDateString('en-GB',{weekday:'short',day:'2-digit',month:'short',timeZone:'UTC'})};
  });
}
export function filterPortalWip(cards,tab,query='') {
  const search=query.trim().toLowerCase();
  return cards.filter(c=>c.tab===tab && (!search || [c.orderNumber,c.company,c.description,c.salesperson,c.productionLocation].some(v=>String(v||'').toLowerCase().includes(search))));
}
export function isPortalWrite(input, options={}) {
  const url=typeof input==='string'?input:input?.url;
  const method=String(options.method || input?.method || 'GET').toUpperCase();
  return !!url && /\/api\//.test(url) && !['GET','HEAD','OPTIONS'].includes(method)
    && !/\/auth\/|\/enrich|\/pull|\/query|\/lookup|\/preview|\/generate|\/calculate|\/compare/.test(url);
}
