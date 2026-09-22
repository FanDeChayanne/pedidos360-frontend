import { ComponentFixture, TestBed } from '@angular/core/testing';
import { LoginFailedComponent } from './login-failed';

describe('LoginFailed', () => {
  let component: LoginFailedComponent;
  let fixture: ComponentFixture<LoginFailedComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [LoginFailedComponent],
    }).compileComponents();

    fixture = TestBed.createComponent(LoginFailedComponent);
    component = fixture.componentInstance;
    await fixture.whenStable();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
