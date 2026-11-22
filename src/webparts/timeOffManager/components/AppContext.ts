import * as React from 'react';
import { IConfig, ICurrentUser, IEmployee, ILeaveTemplate, Role } from '../models/types';
import { LeaveService } from '../services/LeaveService';
import { SettingsStore } from '../services/SettingsStore';
import { SPClient } from '../services/SPClient';
import { Engine, IWorld } from '../services/Engine';
import { GraphService } from '../services/GraphService';
import { IDirectory } from '../services/workflow';
import { ILicenseStatus } from '../services/licensing/license';

export interface IRoute {
  path: string;
  parts: string[];
  query: Record<string, string>;
}

export interface IHostInfo {
  inTeams: boolean;
  /** What the web part shows: the whole app, or only the team calendar / my leave (page layouts). */
  view: 'app' | 'calendar' | 'my';
}

export type ToastKind = 'success' | 'error' | 'info' | 'warning';

export interface IRequestFormOptions {
  employeeEmail?: string;
  kind?: 'leave' | 'compoff';
  template?: ILeaveTemplate;
  start?: string;
  leaveType?: string;
}

export interface IAppContext {
  sp: SPClient;
  svc: LeaveService;
  settings: SettingsStore;
  engine: Engine;
  graph?: GraphService;
  config: IConfig;
  me: ICurrentUser;
  myEmployee?: IEmployee;
  role: Role;
  isAdmin: boolean;
  /** HR or admin. */
  isHr: boolean;
  /** Has people reporting to them, or approves requests. */
  isManager: boolean;
  world: IWorld;
  dir: IDirectory;
  /** The current leave year. */
  year: number;
  license: ILicenseStatus;
  /** Employees within the licence's user count (lower-case emails). */
  licensed: Set<string>;
  branding: { logo: string; custom: boolean };
  tenantId?: string;
  appVersion: string;
  host: IHostInfo;
  pageUrl: string;
  route: IRoute;
  navigate: (path: string) => void;
  toast: (message: string, kind?: ToastKind) => void;
  reload: () => Promise<void>;
  saveConfig: (config: IConfig, what?: string) => Promise<void>;
  saveLogo: (dataUrl: string) => Promise<void>;
  refreshLicense: () => Promise<void>;
  openRequest: (o?: IRequestFormOptions) => void;
  notificationsTick: number;
  bumpNotifications: () => void;
}

export const AppCtx = React.createContext<IAppContext | undefined>(undefined);

export function useApp(): IAppContext {
  const ctx = React.useContext(AppCtx);
  if (!ctx) throw new Error('AppCtx missing');
  return ctx;
}

export function parseRoute(hash: string): IRoute {
  const raw = (hash || '').replace(/^#\/?/, '');
  const [path, qs] = raw.split('?');
  const query: Record<string, string> = {};
  (qs || '').split('&').forEach((pair) => {
    if (!pair) return;
    const [k, v] = pair.split('=');
    query[decodeURIComponent(k)] = decodeURIComponent(v || '');
  });
  const clean = (path || '').replace(/\/$/, '');
  return { path: clean, parts: clean ? clean.split('/') : [], query };
}

export function errorText(e: unknown): string {
  if (e instanceof Error) return e.message;
  return String(e);
}
