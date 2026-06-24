export interface ProductInfo {
  codigo: number;
  nome: string;
  custo_anter: number;
  custo_atual: number;
  variacao: number;
}

export interface ProcessedData {
  [loja: string]: {
    [nota: string]: {
      [fornecedor: string]: ProductInfo[];
    };
  };
}

export interface UserSession {
  username: string;
  isAuthenticated: boolean;
}
