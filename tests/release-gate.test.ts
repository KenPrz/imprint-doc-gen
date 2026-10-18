import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, writeFile, readFile, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { spawnSync } from 'node:child_process';
test('container publication gate fails closed until explicit acceptance is complete',async()=>{
 const folder=await mkdtemp(join(tmpdir(),'imprint-release-gate-'));
 try {
  await mkdir(join(folder,'scripts','assurance'),{recursive:true});await mkdir(join(folder,'docs'));
  const script=join(folder,'scripts','assurance','release-gate.mjs');await writeFile(script,await readFile('scripts/assurance/release-gate.mjs'));
  for(const plan of [{},{releaseApproved:false,remainingAcceptance:[]},{releaseApproved:'true',remainingAcceptance:[]},{releaseApproved:true,remainingAcceptance:['customer baselines']},{releaseApproved:true,remainingAcceptance:null}]){
   await writeFile(join(folder,'docs','assurance-plan.json'),JSON.stringify(plan));assert.equal(spawnSync(process.execPath,[script]).status,1);
  }
  await writeFile(join(folder,'docs','assurance-plan.json'),JSON.stringify({releaseApproved:true,remainingAcceptance:[]}));assert.equal(spawnSync(process.execPath,[script]).status,0);
 }finally{await rm(folder,{recursive:true,force:true})}
});
