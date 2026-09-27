import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';
import { Producto } from './catalogo';
import { environment } from '../../environments/environment';

@Injectable({
  providedIn: 'root'
})
export class CatalogoService {

  private readonly apiUrl = `${environment.apiConfig.uri}/api/catalog/productos`;

  constructor(private http: HttpClient) {}

  listarProductos(): Observable<Producto[]> {
    return this.http.get<Producto[]>(this.apiUrl);
  }

  crearProducto(producto: Omit<Producto, 'id'>): Observable<Producto> {
    return this.http.post<Producto>(
      this.apiUrl,
      producto
    );
  }

  aumentarStock(id: number, cantidad: number): Observable<Producto> {
    return this.http.put<Producto>(
      `${this.apiUrl}/${id}/stock/aumentar?cantidad=${cantidad}`,
      {}
    );
  }

}