import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';
import { environment } from '../../environments/environment';

@Injectable({
  providedIn: 'root'
})
export class PedidoService {

  private readonly apiUrl = `${environment.apiConfig.uri}/api/pedidos`;

  constructor(private http: HttpClient) {}

  listarPedidos(): Observable<any[]> {
    return this.http.get<any[]>(this.apiUrl);
  }
}