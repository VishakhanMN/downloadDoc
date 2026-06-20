import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ViewDocument } from './view-document';

describe('ViewDocument', () => {
  let component: ViewDocument;
  let fixture: ComponentFixture<ViewDocument>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ViewDocument]
    })
    .compileComponents();

    fixture = TestBed.createComponent(ViewDocument);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
