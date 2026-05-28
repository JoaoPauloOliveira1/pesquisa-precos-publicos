import pg from "pg";
import { env } from "../config/env.js";

const { Pool } = pg;

export const pool = new Pool({
  connectionString: env.databaseUrl,
  ssl: env.databaseUrl ? { rejectUnauthorized: false } : false,
});

export async function consultarUmaLinha(texto, params = []) {
  const { rows } = await pool.query(texto, params);
  return rows[0] || null;
}

export async function consultarVariasLinhas(texto, params = []) {
  const { rows } = await pool.query(texto, params);
  return rows;
}
