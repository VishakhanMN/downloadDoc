import { Routes } from '@angular/router';
import { DocumentComponent } from './document/document';

export const routes: Routes = [{
    path: '',
    loadComponent: () => import('./document/document').then(m => m.DocumentComponent)
}];
