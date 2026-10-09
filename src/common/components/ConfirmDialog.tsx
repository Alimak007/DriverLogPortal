/**
 * Confirmation is ALWAYS a Fluent dialog, never a browser confirm box.
 *
 * Destructive confirmations NAME EVERY CONSEQUENCE before they act: who gets
 * notified, what is released, what cannot be undone. A dialog that just says
 * "Are you sure?" is not acceptable here.
 *
 * Imported by path, not from the package root, so the bundle does not pull in
 * the whole of Fluent. The budget is 500 KB gzipped per web part.
 */

import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import * as React from 'react';
import { Button, FormRow, Select, TextArea } from './Shared';
import styles from './Shared.module.scss';

export interface IConfirmDialogProps {
  open: boolean;
  title: string;
  /** What is about to happen, in one line. */
  message: string;
  /** Every consequence, one per line. Shown as a list. */
  consequences?: string[];
  confirmLabel: string;
  destructive?: boolean;
  /** When set, a reason is required before Confirm is enabled. */
  requireReason?: boolean;
  reasonLabel?: string;
  /** When given, the reason is a dropdown rather than free text. */
  reasonOptions?: string[];
  busy?: boolean;
  onConfirm: (reason?: string) => void;
  onCancel: () => void;
}

export const ConfirmDialog: React.FC<IConfirmDialogProps> = ({
  open,
  title,
  message,
  consequences,
  confirmLabel,
  destructive,
  requireReason,
  reasonLabel = 'Reason',
  reasonOptions,
  busy,
  onConfirm,
  onCancel,
}) => {
  const [reason, setReason] = React.useState('');
  const [notes, setNotes] = React.useState('');
  const [touched, setTouched] = React.useState(false);

  React.useEffect(() => {
    if (open) {
      setReason('');
      setNotes('');
      setTouched(false);
    }
  }, [open]);

  const needsNotes = reasonOptions !== undefined && reason === 'Other';
  const value = needsNotes ? `${reason}: ${notes}` : reason;
  const missing = requireReason === true && value.trim().length === 0;
  const notesMissing = needsNotes && notes.trim().length === 0;
  const blocked = missing || notesMissing;

  return (
    <Dialog
      hidden={!open}
      onDismiss={onCancel}
      dialogContentProps={{
        type: DialogType.normal,
        title,
        subText: message,
      }}
      modalProps={{ isBlocking: true }}
    >
      {consequences && consequences.length > 0 ? (
        <ul style={{ margin: '0 0 12px 0', paddingLeft: 20 }}>
          {consequences.map((c) => (
            <li key={c} className={styles.hint} style={{ marginBottom: 4 }}>
              {c}
            </li>
          ))}
        </ul>
      ) : undefined}

      {requireReason ? (
        <>
          <FormRow
            label={reasonLabel}
            required
            error={touched && missing ? 'A reason is required.' : undefined}
          >
            {reasonOptions ? (
              <Select
                value={reason}
                onChange={(e) => {
                  setReason(e.target.value);
                  setTouched(true);
                }}
              >
                <option value="">Choose a reason</option>
                {reasonOptions.map((o) => (
                  <option key={o} value={o}>
                    {o}
                  </option>
                ))}
              </Select>
            ) : (
              <TextArea
                value={reason}
                onChange={(e) => {
                  setReason(e.target.value);
                  setTouched(true);
                }}
                onBlur={() => setTouched(true)}
              />
            )}
          </FormRow>

          {needsNotes ? (
            <FormRow
              label="Please explain"
              required
              error={touched && notesMissing ? 'A note is required when the reason is Other.' : undefined}
            >
              <TextArea
                value={notes}
                onChange={(e) => {
                  setNotes(e.target.value);
                  setTouched(true);
                }}
              />
            </FormRow>
          ) : undefined}
        </>
      ) : undefined}

      <DialogFooter>
        <Button kind="secondary" onClick={onCancel} disabled={busy}>
          Cancel
        </Button>
        <Button
          kind={destructive ? 'destructive' : 'primary'}
          onClick={() => {
            setTouched(true);
            if (!blocked) onConfirm(requireReason ? value : undefined);
          }}
          disabled={busy || blocked}
        >
          {busy ? 'Working…' : confirmLabel}
        </Button>
      </DialogFooter>
    </Dialog>
  );
};
