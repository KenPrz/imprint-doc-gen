import { createApp } from './app.ts';
import { configuredStore } from './config-store.ts';
const store=await configuredStore();
if(store&&process.env.IMPRINT_DATA_DIR)throw new Error('Choose DATABASE_URL or IMPRINT_DATA_DIR, not both');
const app=createApp({apiKey:process.env.IMPRINT_API_KEY??'',converterUrl:process.env.GOTENBERG_URL,logger:true,dataDir:process.env.IMPRINT_DATA_DIR,store,runQueue:!store,concurrency:Number(process.env.IMPRINT_CONCURRENCY??2)});
await app.listen({host:process.env.HOST??'127.0.0.1',port:Number(process.env.PORT??3000)});
for(const signal of ['SIGTERM','SIGINT'])process.once(signal,()=>{void app.close().then(()=>process.exit(0));});
