import * as React from 'react';
import { Callout, DirectionalHint, Icon, IconButton } from '@fluentui/react';
import { INotification } from '../../models/types';
import { useApp } from '../AppContext';
import { relativeTime } from '../../services/dates';

const FONT = "'DM Sans', 'Segoe UI', sans-serif";

function iconFor(title: string): string {
  if (/approved|credited/i.test(title)) return 'CheckMark';
  if (/declined/i.test(title)) return 'Cancel';
  if (/cancel/i.test(title)) return 'EventDeclined';
  if (/reminder|escalated/i.test(title)) return 'Clock';
  return 'Inbox';
}

/** Top-bar bell: the signed-in person's notifications (BrillTM_Notifications), refreshed every minute. */
export function NotificationBell(): JSX.Element | null {
  const { svc, me, navigate, config, notificationsTick } = useApp();
  const [items, setItems] = React.useState<INotification[]>([]);
  const [open, setOpen] = React.useState(false);
  const anchor = React.useRef<HTMLDivElement>(null);

  const load = React.useCallback(() => {
    svc
      .loadNotifications(me.email)
      .then(setItems)
      .catch(() => undefined);
  }, [svc, me.email]);

  React.useEffect(() => {
    if (!config.notifications.inApp) return undefined;
    load();
    const t = window.setInterval(() => !document.hidden && load(), 60000);
    return () => window.clearInterval(t);
  }, [load, notificationsTick, config.notifications.inApp]);

  if (!config.notifications.inApp) return null;
  const unread = items.filter((n) => !n.read);
  const markRead = (list: INotification[]): void => {
    setItems((cur) => cur.map((n) => (list.some((x) => x.id === n.id) ? { ...n, read: true } : n)));
    list.forEach((n) => void svc.markRead(n.id).catch(() => undefined));
  };

  return (
    <div ref={anchor} style={{ position: 'relative' }}>
      <IconButton iconProps={{ iconName: unread.length ? 'RingerSolid' : 'Ringer' }} onClick={() => setOpen(!open)} ariaLabel={`Notifications, ${unread.length} unread`} />
      {unread.length > 0 && <span className="to-bell-count">{unread.length > 9 ? '9+' : unread.length}</span>}
      {open && anchor.current && (
        <Callout target={anchor.current} onDismiss={() => setOpen(false)} directionalHint={DirectionalHint.bottomRightEdge} isBeakVisible={false} styles={{ calloutMain: { width: 380, maxHeight: 480, borderRadius: 10 } }}>
          <div style={{ padding: '12px 16px', display: 'flex', alignItems: 'center', borderBottom: '1px solid #dfe5e6', fontFamily: FONT }}>
            <strong style={{ fontFamily: "'Manrope', 'Segoe UI', sans-serif" }}>Notifications</strong>
            <span style={{ flex: 1 }} />
            {unread.length > 0 && (
              <button style={{ border: 0, background: 'none', cursor: 'pointer', color: '#17232e', fontWeight: 600, textDecoration: 'underline', font: 'inherit', fontSize: 12 }} onClick={() => markRead(unread)}>
                Mark all read
              </button>
            )}
          </div>
          {!items.length && <div style={{ padding: 24, textAlign: 'center', color: '#647079', fontFamily: FONT }}>You’re all caught up.</div>}
          {items.map((n) => (
            <div
              key={n.id}
              role="button"
              tabIndex={0}
              onClick={() => {
                setOpen(false);
                if (!n.read) markRead([n]);
                if (n.link) navigate(n.link);
              }}
              style={{ display: 'flex', gap: 12, padding: '11px 16px', cursor: 'pointer', background: n.read ? undefined : '#f3f8ec', borderBottom: '1px solid #eef1f1', fontFamily: FONT }}
            >
              <span style={{ width: 30, height: 30, borderRadius: 8, background: '#e7efde', display: 'grid', placeItems: 'center', flex: 'none' }}>
                <Icon iconName={iconFor(n.title)} style={{ fontSize: 14, color: '#17232e' }} />
              </span>
              <div style={{ minWidth: 0 }}>
                <div style={{ fontWeight: n.read ? 500 : 700, fontSize: 13, color: '#17232e' }}>{n.title}</div>
                <div style={{ fontSize: 12, color: '#647079' }}>{n.body}</div>
                <div style={{ fontSize: 11, color: '#8a959c', marginTop: 2 }}>{relativeTime(n.created)}</div>
              </div>
            </div>
          ))}
        </Callout>
      )}
    </div>
  );
}
