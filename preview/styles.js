// Preview stand-in for the SPFx CSS-module loader: injects the compiled stylesheet and exposes its class names.
import css from '../lib/webparts/timeOffManager/components/TimeOff.module.scss.css';

const style = document.createElement('style');
style.textContent = css.replace(/:global\s*/g, '');
document.head.appendChild(style);

export default { root: 'root' };
