import { Injectable } from "@angular/core";
import { HttpClient } from "@angular/common/http";
import { Observable } from "rxjs";
import { Pedido, CrearPedido } from "./pedido";

@Injectable({
  providedIn: "root",
})
export class PedidoService {
  private readonly apiUrl = "http://localhost:8080/api/pedidos";

  constructor(private http: HttpClient) {}

  listarPedidos(): Observable<Pedido[]> {
    return this.http.get<Pedido[]>(this.apiUrl);
  }

  crearPedido(pedido: CrearPedido): Observable<Pedido> {
    return this.http.post<Pedido>(this.apiUrl, pedido);
  }

  cambiarEstado(id: number, estado: string): Observable<Pedido> {
    return this.http.put<Pedido>(
      `${this.apiUrl}/${id}/estado?estado=${estado}`,
      {},
    );
  }
}
