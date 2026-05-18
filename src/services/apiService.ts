import { ApiResponseItem } from '../types';

const API_URL = (import.meta as any).env.VITE_API_URL;

function checkConfig() {
  if (!API_URL) {
    throw new Error("VITE_API_URL não configurado nas variáveis de ambiente.");
  }
}

export async function postSqlQuery<T = any[]>(query: string, id_score: string): Promise<T> {
  checkConfig();

  const response = await fetch(API_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ query, id_score })
  });

  if (!response.ok) {
    throw new Error(`Erro: ${response.status} ao acessar ${API_URL}`);
  }

  return response.json();
}

// Fixed constant queries (provided by user)
export const TEST_API_QUERY = "SELECT * FROM dual"; 
export const buildMontagemKitProducaoQuery = () => "SELECT * FROM montagem_kit_producao";
export const buildRecebimentoResumoQuery = () => "SELECT * FROM recebimento_resumo";
export const buildCatalogacaoResumoQuery = () => "SELECT * FROM catalogacao_resumo";
export const buildCatalogacaoPendenteQuery = () => "SELECT * FROM catalogacao_pendente";
export const CATALOGACAO_BUFFER_QUERY = "SELECT * FROM catalogacao_buffer";
export const CATALOGACAO_AGING_QUERY = "SELECT * FROM catalogacao_aging";
export const MONTAGEM_KIT_COMPONENTS_STOCK_QUERY = "SELECT * FROM montagem_kit_componentes_estoque";
export const MONTAGEM_KIT_REPLENISHMENT_QUERY = "SELECT * FROM montagem_kit_reabastecimento";
export const CARGA_APP_TRIAGEM_QUERY = "SELECT * FROM carga_app_triagem";
export const CARGA_APP_TRIAGEM_STOCK_QUERY = "SELECT * FROM carga_app_triagem_estoque";
export const SALA_BATERIAS_REALIZADO_QUERY = "SELECT * FROM sala_baterias_realizado";
export const SALA_BATERIAS_STOCK_QUERY = "SELECT * FROM sala_baterias_stock";
export const EXPEDICAO_REALIZADO_QUERY = "SELECT * FROM expedicao_realizado";
export const SEPARACAO_COMPONENTES_STOCK_QUERY = "SELECT * FROM separacao_componentes_stock";
export const SEPARACAO_COMPONENTES_REALIZADO_QUERY = "SELECT * FROM separacao_componentes_realizado";

export async function fetchApiData(): Promise<ApiResponseItem[]> {
  try {
    return await postSqlQuery<ApiResponseItem[]>(TEST_API_QUERY, "recebimentos_rdts");
  } catch (error) {
    console.error("Fetch error:", error);
    throw error;
  }
}
