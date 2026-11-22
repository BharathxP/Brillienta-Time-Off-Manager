import * as React from 'react';
import { Icon, IconButton, PrimaryButton, TooltipHost } from '@fluentui/react';
import { IRequestFormOptions, useApp } from './AppContext';
import { Avatar } from './common/Ui';
import { NotificationBell } from './common/NotificationBell';
import { waitingFor } from '../services/workflow';
import { describeStatus } from '../services/licensing/license';
import { Home } from './home/Home';
import { MyRequests } from './requests/MyRequests';
import { RequestForm } from './requests/RequestForm';
import { Approvals } from './approvals/Approvals';
import { TeamCalendar } from './calendar/TeamCalendar';
import { Reports } from './dashboard/Reports';
import { EmployeesPage } from './admin/EmployeesPage';
import { MyTeam } from './team/MyTeam';
import { AdminPage } from './admin/AdminPage';

interface INavItem {
  key: string;
  label: string;
  icon: string;
  section?: string;
  count?: number;
}

const ROLE_LABEL: Record<string, string> = { admin: 'Admin', hr: 'HR', manager: 'Manager', employee: 'Employee' };

export function Shell(props: { requestForm?: IRequestFormOptions; closeRequestForm: () => void; siteTitle: string }): JSX.Element {
  const app = useApp();
  const { route, navigate, isHr, isAdmin, isManager, config, me, world, role, license, host } = app;
  const [navOpen, setNavOpen] = React.useState(false);
  const [collapsed, setCollapsed] = React.useState(false);
  const waiting = waitingFor(world.requests, me.email).length;
  const items: INavItem[] = [
    { key: 'home', label: 'Overview', icon: 'ViewDashboard', section: 'Your time off' },
    { key: 'requests', label: 'My leave', icon: 'Calendar' },
    { key: 'calendar', label: 'Team calendar', icon: 'CalendarWeek' }
  ];
  const hasReports = world.employees.some((e) => e.active && (e.managerEmail || '').toLowerCase() === me.email.toLowerCase());
  if (isManager || isHr || waiting) items.push({ key: 'approvals', label: 'Approvals', icon: 'CheckList', section: 'Approve', count: waiting });
  if (hasReports) items.push({ key: 'team', label: 'My team', icon: 'Group', section: isManager || isHr || waiting ? undefined : 'Approve' });
  if (isHr) {
    items.push({ key: 'people', label: 'Employees & balances', icon: 'People', section: 'Manage' });
    items.push({ key: 'reports', label: 'Reports', icon: 'BarChartVertical' });
    items.push({ key: 'admin', label: isAdmin ? 'Admin' : 'HR settings', icon: 'Settings' });
  }
  const fixedView = host.view === 'calendar' ? 'calendar' : host.view === 'my' ? 'requests' : '';
  const path = fixedView && !route.parts[0] ? fixedView : route.parts[0] || 'home';
  const allowed = items.some((i) => i.key === path) ? path : 'home';
  let page: JSX.Element;
  switch (allowed) {
    case 'requests':
      page = <MyRequests />;
      break;
    case 'calendar':
      page = <TeamCalendar />;
      break;
    case 'approvals':
      page = <Approvals />;
      break;
    case 'team':
      page = <MyTeam />;
      break;
    case 'people':
      page = <EmployeesPage />;
      break;
    case 'reports':
      page = <Reports />;
      break;
    case 'admin':
      page = <AdminPage section={route.parts[1]} />;
      break;
    default:
      page = <Home />;
  }
  const current = items.filter((i) => i.key === allowed)[0];
  let lastSection: string | undefined;
  const showTrial = license.state === 'trial' || license.state === 'grace';

  return (
    <div className="to-app">
      <nav className={'to-nav' + (navOpen ? ' is-open' : '') + (collapsed ? ' is-collapsed' : '')} aria-label={config.appName}>
        <div className="to-nav-brand">
          <img className="to-logo" src={app.branding.logo} alt="" />
          <span className="to-nav-label">{config.appName}</span>
          <button type="button" className="to-nav-collapse" onClick={() => setCollapsed(!collapsed)} aria-label={collapsed ? 'Expand navigation' : 'Collapse navigation'}>
            <span aria-hidden="true">{collapsed ? '»' : '«'}</span>
          </button>
        </div>
        <div className="to-nav-workspace" title={props.siteTitle}>
          <strong>{config.orgName || props.siteTitle}</strong>
          <span>{ROLE_LABEL[role]}</span>
        </div>
        <div className="to-nav-cta">
          <PrimaryButton iconProps={{ iconName: 'Add' }} text="Request time off" onClick={() => app.openRequest()} styles={{ root: { width: '100%' } }} />
        </div>
        {items.map((it) => {
          const header = it.section && it.section !== lastSection ? it.section : undefined;
          if (it.section) lastSection = it.section;
          return (
            <React.Fragment key={it.key}>
              {header && (
                <div className="to-nav-section">
                  <span className="to-nav-label">{header}</span>
                </div>
              )}
              <button
                className={'to-nav-item' + (allowed === it.key ? ' is-active' : '')}
                aria-current={allowed === it.key ? 'page' : undefined}
                title={collapsed ? it.label : undefined}
                onClick={() => {
                  setNavOpen(false);
                  navigate(it.key);
                }}
              >
                <Icon iconName={it.icon} />
                <span className="to-nav-label">{it.label}</span>
                {!!it.count && <span className="to-nav-count">{it.count}</span>}
              </button>
            </React.Fragment>
          );
        })}
        <div className="to-nav-footer" title={`Version ${app.appVersion}`}>
          <span className="to-msicon">
            <i />
            <i />
            <i />
            <i />
          </span>
          <div className="to-nav-label">
            Brillienta
            <span className="to-sub">
              {world.employees.filter((e) => e.active).length} people · v{app.appVersion}
            </span>
          </div>
        </div>
      </nav>
      <div className="to-main" data-to-main="1">
        <div className="to-topbar">
          <IconButton className="to-menu-btn" iconProps={{ iconName: 'GlobalNavButton' }} onClick={() => setNavOpen(!navOpen)} ariaLabel="Menu" />
          <span className="to-crumb">
            Time off<span className="to-crumb-sep">/</span>
            <strong>{current ? current.label : 'Overview'}</strong>
          </span>
          <span className="to-spacer" />
          {showTrial && (
            <button className="to-trial-pill" onClick={() => isAdmin && navigate('admin/licence')} title={describeStatus(license)}>
              <Icon iconName="Clock" /> {describeStatus(license)}
            </button>
          )}
          <div className="to-topbar-actions">
            {(isManager || isHr) && (
              <TooltipHost content="Waiting for your approval">
                <button className="to-iconbtn" style={{ width: 36, height: 36 }} onClick={() => navigate('approvals')} aria-label={`Approvals, ${waiting} waiting`}>
                  <Icon iconName="CheckList" />
                  {waiting > 0 && <span className="to-bell-count">{waiting}</span>}
                </button>
              </TooltipHost>
            )}
            <NotificationBell />
            <button className="to-iconbtn" style={{ width: 36, height: 36 }} onClick={() => navigate('requests')} aria-label="My leave" title={`${me.title} — ${ROLE_LABEL[role]}`}>
              <Avatar email={me.email} name={me.title} />
            </button>
          </div>
        </div>
        <div className="to-page" key={route.path} onClick={() => navOpen && setNavOpen(false)}>
          {page}
        </div>
      </div>
      {props.requestForm && <RequestForm options={props.requestForm} onClose={props.closeRequestForm} />}
    </div>
  );
}
