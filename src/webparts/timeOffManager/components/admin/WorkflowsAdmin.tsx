import * as React from 'react';
import { DefaultButton, Dropdown, Icon, IconButton, PrimaryButton, TextField } from '@fluentui/react';
import { ApproverKind, IApprovalLevel, IWorkflow } from '../../models/types';
import { ISectionProps } from './AdminPage';
import { Card, PeopleField, NumberBox } from '../common/Ui';

const KINDS: Array<{ key: ApproverKind; text: string }> = [
  { key: 'manager', text: 'Manager' },
  { key: 'skipManager', text: "Manager's manager" },
  { key: 'departmentHead', text: 'Department head' },
  { key: 'hr', text: 'HR (any HR person)' },
  { key: 'person', text: 'A named person' }
];

/** Multi-level approval workflows: levels in order, each optionally only for longer requests. */
export function WorkflowsAdmin(props: ISectionProps): JSX.Element {
  const { draft, setDraft } = props;
  const update = (w: IWorkflow): void => setDraft({ ...draft, workflows: draft.workflows.map((x) => (x.key === w.key ? w : x)) });
  const add = (): void => {
    let key = 'workflow';
    for (let i = 1; draft.workflows.some((w) => w.key === key); i++) key = `workflow${i}`;
    setDraft({ ...draft, workflows: draft.workflows.concat({ key, name: 'New workflow', levels: [{ kind: 'manager', label: 'Manager', minDays: 0 }] }) });
  };
  return (
    <>
      <Card title="How approval works" icon="Info">
        <p className="to-muted" style={{ margin: 0 }}>
          Each leave type uses one workflow (Leave types › Approval). A request goes through the levels in order; a level with “from N days” only applies to requests of at least that length. The requester never approves their own leave;
          a level without an approver (e.g. no manager) goes to HR; a level with the same approver as the one before is merged. Delegations apply to every level. Auto-approval rules on the leave type skip the workflow.
        </p>
      </Card>
      {draft.workflows.map((w) => {
        const usedBy = draft.leaveTypes.filter((t) => t.workflowKey === w.key);
        const setLevel = (i: number, patch: Partial<IApprovalLevel>): void => update({ ...w, levels: w.levels.map((l, j) => (j === i ? { ...l, ...patch } : l)) });
        return (
          <Card
            key={w.key}
            title={w.name}
            icon="Flow"
            actions={
              <DefaultButton
                text="Delete"
                disabled={usedBy.length > 0 || draft.workflows.length < 2}
                title={usedBy.length ? 'Used by leave types' : undefined}
                onClick={() => setDraft({ ...draft, workflows: draft.workflows.filter((x) => x.key !== w.key) })}
              />
            }
          >
            <TextField label="Name" value={w.name} onChange={(_, v) => update({ ...w, name: v || '' })} styles={{ root: { maxWidth: 420 } }} />
            <div className="to-levels">
              {w.levels.map((l, i) => (
                <div key={i} className="to-level">
                  <span className="to-flow-num">{i + 1}</span>
                  <div className="to-level-body">
                    <div className="to-form-grid">
                      <Dropdown label="Approver" selectedKey={l.kind} options={KINDS} onChange={(_, o) => o && setLevel(i, { kind: o.key as ApproverKind, label: l.label || o.text })} />
                      <TextField label="Label" value={l.label} onChange={(_, v) => setLevel(i, { label: v || '' })} />
                      <NumberBox label="Only from (days, 0 = always)" min={0} max={365} step={0.5} value={String(l.minDays)} onChange={(_, v) => setLevel(i, { minDays: Number(v) || 0 })} />
                      {l.kind === 'person' && <PeopleField single label="Person" value={l.person ? [l.person] : []} onChange={(v) => setLevel(i, { person: v[0] })} />}
                    </div>
                  </div>
                  <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Remove level" disabled={w.levels.length < 2} onClick={() => update({ ...w, levels: w.levels.filter((_, j) => j !== i) })} />
                </div>
              ))}
            </div>
            <div className="to-row" style={{ marginTop: 10 }}>
              <DefaultButton iconProps={{ iconName: 'Add' }} text="Add level" onClick={() => update({ ...w, levels: w.levels.concat({ kind: 'hr', label: 'HR', minDays: 0 }) })} />
              <span className="to-small to-muted">
                <Icon iconName="Airplane" /> Used by: {usedBy.map((t) => t.name).join(', ') || 'no leave type'}
              </span>
            </div>
          </Card>
        );
      })}
      <div>
        <PrimaryButton iconProps={{ iconName: 'Add' }} text="New workflow" onClick={add} />
      </div>
    </>
  );
}
