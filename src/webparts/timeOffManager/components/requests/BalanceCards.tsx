import * as React from 'react';
import { Icon, TooltipHost } from '@fluentui/react';
import { useApp } from '../AppContext';
import { IEmployee } from '../../models/types';
import { IBalance } from '../../services/balances';
import { fmtNum } from '../../services/dates';

/** How a balance is made up, line by line (tooltip and detail views). */
export function balanceLines(b: IBalance): Array<[string, string]> {
  if (b.unlimited) return [['Policy', 'No limit'], ['Taken', fmtNum(b.used)], ['Pending', fmtNum(b.pending)]];
  const out: Array<[string, string]> = [];
  if (b.earned || b.expired) {
    out.push(['Earned', fmtNum(b.earned)]);
    if (b.expired) out.push(['Lapsed', '−' + fmtNum(b.expired)]);
  } else {
    out.push([b.annual !== b.entitled ? `Accrued so far (of ${fmtNum(b.annual)})` : 'Entitlement', fmtNum(b.entitled)]);
    if (b.tenureExtra) out.push(['…of which for length of service', fmtNum(b.tenureExtra)]);
    if (b.carried) out.push(['Carried forward', fmtNum(b.carried)]);
    if (b.adjusted) out.push(['Adjustments', (b.adjusted > 0 ? '+' : '') + fmtNum(b.adjusted)]);
  }
  out.push(['Taken', '−' + fmtNum(b.used)]);
  if (b.pending) out.push(['Pending', '−' + fmtNum(b.pending)]);
  out.push(['Available', fmtNum(b.available)]);
  return out;
}

export function BalanceCards(props: { employee?: IEmployee; year: number; onRequest?: (type: string) => void; compact?: boolean; showAll?: boolean; limit?: number }): JSX.Element {
  const { engine, world, config } = useApp();
  const [expanded, setExpanded] = React.useState(false);
  const balances = engine.balancesFor(world, props.employee, props.year).filter((b) => b.eligible && (props.showAll || !b.unlimited || b.used > 0 || b.pending > 0));
  const relevant = props.showAll ? balances : balances.filter((b) => (b.unlimited ? b.used + b.pending > 0 : b.annual > 0 || b.available !== 0 || b.used > 0 || b.earned > 0));
  // With a limit: the first types in the admin's order, plus any the person has used, pending or earned.
  const limited = props.limit && !expanded ? relevant.filter((b, i) => i < (props.limit as number) || b.used > 0 || b.pending > 0 || b.earned > 0 || b.carried > 0) : relevant;
  const shown = limited;
  return (
    <>
    <div className={'to-balances' + (props.compact ? ' is-compact' : '')}>
      {shown.map((b) => {
        const t = config.leaveTypes.filter((x) => x.key === b.leaveType)[0];
        const total = b.unlimited ? 0 : Math.max(b.entitled + b.carried + b.adjusted + b.earned - b.expired, b.used + b.pending + Math.max(0, b.available));
        const pct = (n: number): string => (total > 0 ? Math.max(0, Math.min(100, (n / total) * 100)) : 0) + '%';
        return (
          <TooltipHost
            key={b.leaveType}
            content={
              <div style={{ padding: 4, minWidth: 200 }}>
                {balanceLines(b).map(([k, v]) => (
                  <div key={k} style={{ display: 'flex', justifyContent: 'space-between', gap: 12 }}>
                    <span>{k}</span>
                    <b>{v}</b>
                  </div>
                ))}
              </div>
            }
          >
            <button type="button" className="to-balance" style={{ '--to-type': t.color } as unknown as React.CSSProperties} onClick={() => props.onRequest && props.onRequest(b.leaveType)}>
              <span className="to-balance-head">
                <span className="to-balance-icon">
                  <Icon iconName={t.icon} />
                </span>
                <span className="to-balance-name">{t.name}</span>
              </span>
              <span className="to-balance-value">
                {b.unlimited ? '∞' : fmtNum(b.available)}
                <small>{b.unlimited ? `${fmtNum(b.used)} taken` : 'days left'}</small>
              </span>
              {!b.unlimited && (
                <span className="to-balance-bar" aria-hidden="true">
                  <i className="is-used" style={{ width: pct(b.used) }} />
                  <i className="is-pending" style={{ width: pct(b.pending) }} />
                </span>
              )}
              <span className="to-balance-foot">
                {b.unlimited ? 'No limit' : `${fmtNum(b.used)} taken${b.pending ? ` · ${fmtNum(b.pending)} pending` : ''}${t.entitlement.mode === 'accrual' ? ` · accrues ${t.entitlement.accrual.frequency}` : ''}${b.expired ? ` · ${fmtNum(b.expired)} lapsed` : ''}`}
              </span>
            </button>
          </TooltipHost>
        );
      })}
    </div>
    {props.limit && relevant.length > limited.length && (
      <button type="button" className="to-link" style={{ marginTop: 10 }} onClick={() => setExpanded(true)}>
        Show all {relevant.length} leave types
      </button>
    )}
    {props.limit && expanded && (
      <button type="button" className="to-link" style={{ marginTop: 10 }} onClick={() => setExpanded(false)}>
        Show fewer
      </button>
    )}
    </>
  );
}
