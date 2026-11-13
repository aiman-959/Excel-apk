import { Dialog } from '../ui/Chrome';
import { ui } from '../ui/store';
export default function Stub() {
  return <Dialog title="Coming soon" onClose={() => ui.close()}>…</Dialog>;
}
