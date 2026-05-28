import { criarApp } from "./app.js";
import { assertEnv, env } from "./config/env.js";

assertEnv();

const app = criarApp();

app.listen(env.port, env.host, () => {
  console.log(`SINAPI Node API rodando em http://${env.host}:${env.port}`);
});
