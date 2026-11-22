import * as React from 'react';
import { Checkbox, Icon } from '@fluentui/react';
import { useApp } from '../AppContext';
import { ILeaveRequest, ILeaveType } from '../../models/types';
import { PersonLine, StatusBadge, TypeChip, Empty } from '../common/Ui';
import { fmtRange, relativeTime } from '../../services/dates';
import { lengthText } from './RequestDetails';

/** Requests as a table; a row opens the request. Optional selection for bulk actions. */
export function RequestTable(props: {
  rows: ILeaveRequest[];
  onOpen: (r: ILeaveRequest) => void;
  showPerson?: boolean;
  selected?: number[];
  onSelect?: (ids: number[]) => void;
  empty?: string;
  waitingColumn?: boolean;
}): JSX.Element {
  const { config } = useApp();
  const typeOf = (k: string): ILeaveType | undefined => config.leaveTypes.filter((t) => t.key === k)[0];
  if (!props.rows.length) return <Empty icon="Calendar" title={props.empty || 'Nothing here yet.'} />;
  const sel = props.selected || [];
  const all = props.rows.length > 0 && props.rows.every((r) => sel.indexOf(r.id) >= 0);
  return (
    <div className="to-table-wrap">
      <table className="to-table">
        <thead>
          <tr>
            {props.onSelect && (
              <th style={{ width: 32 }}>
                <Checkbox ariaLabel="Select all" checked={all} onChange={(_, c) => props.onSelect && props.onSelect(c ? props.rows.map((r) => r.id) : [])} />
              </th>
            )}
            {props.showPerson && <th>Employee</th>}
            <th>Type</th>
            <th>When</th>
            <th>Length</th>
            <th>Status</th>
            <th>{props.waitingColumn ? 'Waiting since' : 'Submitted'}</th>
          </tr>
        </thead>
        <tbody>
          {props.rows.map((r) => {
            const step = r.steps[r.currentStep];
            return (
              <tr key={r.id} className="is-clickable" onClick={() => props.onOpen(r)}>
                {props.onSelect && (
                  <td onClick={(e) => e.stopPropagation()}>
                    <Checkbox ariaLabel={`Select ${r.employeeName}`} checked={sel.indexOf(r.id) >= 0} onChange={(_, c) => props.onSelect && props.onSelect(c ? sel.concat(r.id) : sel.filter((x) => x !== r.id))} />
                  </td>
                )}
                {props.showPerson && (
                  <td>
                    <PersonLine email={r.employeeEmail} name={r.employeeName} size="sm" />
                  </td>
                )}
                <td>
                  <TypeChip type={typeOf(r.leaveType)} label={r.kind === 'compoff' ? 'Comp off claim' : undefined} />
                  {r.seriesId && <Icon iconName="RepeatAll" title={r.seriesRule} style={{ marginLeft: 6, color: 'var(--to-muted)' }} />}
                  {r.hasAttachment && <Icon iconName="Attach" title="Has attachments" style={{ marginLeft: 6, color: 'var(--to-muted)' }} />}
                </td>
                <td>{fmtRange(r.start, r.end)}</td>
                <td>{lengthText(r)}</td>
                <td>
                  <StatusBadge status={r.status} />
                  {r.status === 'Pending' && step && <div className="to-small to-muted">{step.label}: {step.approvers.map((a) => a.name).join(', ')}</div>}
                </td>
                <td className="to-small to-muted">{relativeTime(props.waitingColumn && step && step.startedAt ? step.startedAt : r.created)}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}
