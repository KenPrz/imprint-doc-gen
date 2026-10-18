import { readFileSync } from 'node:fs';
const plan=JSON.parse(readFileSync(new URL('../../docs/assurance-plan.json',import.meta.url),'utf8'));
if(plan.releaseApproved!==true || !Array.isArray(plan.remainingAcceptance) || plan.remainingAcceptance.length!==0){
 console.error('Container release blocked: assurance acceptance is incomplete. See docs/assurance-plan.json.');
 process.exitCode=1;
}else console.log('Assurance acceptance completed. Required CI jobs must also pass.');
