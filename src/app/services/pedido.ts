export interface Pedido {
  id: number;
  clienteId: string;
  productoId: number;
  cantidad: number;
  estado: string;
}

export interface CrearPedido {
  clienteId: string;
  productoId: number;
  cantidad: number;
}