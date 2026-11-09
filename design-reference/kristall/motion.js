// One interruptible spring clock drives both the HTML silhouette and its optical
// lens. Different leading/trailing edge responses make a moving drop stretch.
const clamp = (value, low, high) => Math.min(high, Math.max(low, value));

class Spring {
  constructor(value = 0, stiffness = 460, damping = 30) {
    this.value = this.target = value;
    this.velocity = 0;
    this.stiffness = stiffness;
    this.damping = damping;
  }
  snap(value = this.target) { this.value = this.target = value; this.velocity = 0; }
  step(dt) {
    // Small fixed substeps keep pointer interruptions and slow frames stable.
    const count = Math.max(1, Math.ceil(dt / (1 / 180)));
    const step = dt / count;
    for (let i = 0; i < count; i++) {
      this.velocity += ((this.target - this.value) * this.stiffness - this.velocity * this.damping) * step;
      this.value += this.velocity * step;
    }
    if (Math.abs(this.target - this.value) < .012 && Math.abs(this.velocity) < .06) this.snap();
    return this.value !== this.target || this.velocity !== 0;
  }
}

export class LiquidMotion {
  constructor(render) {
    this.render = render;
    this.active = new Set();
    this.frame = 0;
    this.previous = 0;
    this.until = 0;
    this.preference = matchMedia('(prefers-reduced-motion: reduce)');
    this.preference.addEventListener('change', () => {
      if (this.preference.matches) this.finish();
    });
    document.addEventListener('visibilitychange', () => { if (document.hidden) this.finish(); });
  }
  start(controller) {
    this.active.add(controller);
    if (this.preference.matches || document.hidden) {
      // A preference change can precede its change event. Settle the whole
      // clock, including controllers that were already in flight.
      this.finish();
      return;
    }
    this.schedule();
  }
  animateFor(duration = 0) {
    if (document.hidden) { this.finish(); return; }
    this.until = Math.max(this.until, performance.now() + (this.preference.matches ? 0 : duration));
    this.schedule();
  }
  schedule() {
    if (!this.frame) {
      this.previous = performance.now();
      this.frame = requestAnimationFrame(time => this.tick(time));
    }
  }
  tick(time) {
    this.frame = 0;
    if (this.preference.matches || document.hidden) { this.finish(); return; }
    const dt = clamp((time - this.previous) / 1000, 1 / 240, 1 / 30);
    this.previous = time;
    for (const controller of this.active) if (!controller.step(dt)) this.active.delete(controller);
    // Draw in this very frame: never let the refraction trail the DOM silhouette.
    this.render();
    if (this.active.size || time < this.until) this.frame = requestAnimationFrame(next => this.tick(next));
  }
  finish() {
    cancelAnimationFrame(this.frame);
    this.frame = 0;
    this.until = 0;
    for (const controller of this.active) controller.finish();
    this.active.clear();
    this.render();
  }
}

export class LiquidMarker {
  constructor(element, motion, { container = null, radius = 24, clampEdges = true } = {}) {
    this.element = element;
    this.motion = motion;
    this.container = container;
    this.radius = radius;
    this.clampEdges = clampEdges;
    this.left = new Spring();
    this.right = new Spring();
    this.top = new Spring(0, 460, 31);
    this.height = new Spring(0, 460, 31);
    this.pressure = new Spring(1, 510, 25);
    this.initialized = false;
    this.springs = [this.left, this.right, this.top, this.height, this.pressure];
  }
  follow(target, immediate = false) {
    if (!target || !target.offsetWidth) {
      this.element.hidden = true;
      return;
    }
    this.element.hidden = false;
    const parent = this.element.offsetParent;
    const box = target.getBoundingClientRect();
    const origin = parent.getBoundingClientRect();
    const style = getComputedStyle(parent);
    const scaleX = origin.width / parseFloat(style.width);
    const scaleY = origin.height / parseFloat(style.height);
    this.to({ x: (box.x - origin.x) / scaleX, y: (box.y - origin.y) / scaleY, width: box.width / scaleX, height: box.height / scaleY }, immediate);
  }
  to(rect, immediate = false) {
    this.element.hidden = false;
    if (this.initialized && !immediate && Math.abs(rect.x - this.left.target) < .001 && Math.abs(rect.x + rect.width - this.right.target) < .001 && Math.abs(rect.y - this.top.target) < .001 && Math.abs(rect.height - this.height.target) < .001) return;
    const delta = rect.x - this.left.target;
    this.left.stiffness = delta >= 0 ? 370 : 570;
    this.right.stiffness = delta >= 0 ? 570 : 370;
    this.left.damping = delta >= 0 ? 27 : 31;
    this.right.damping = delta >= 0 ? 31 : 27;
    this.left.target = rect.x;
    this.right.target = rect.x + rect.width;
    this.top.target = rect.y;
    this.height.target = rect.height;
    this.pressure.target = 1;
    if (immediate || !this.initialized) {
      this.initialized = true;
      this.finish();
      this.motion.animateFor(0);
    } else {
      if (Math.abs(delta) > .1) this.pressure.velocity -= Math.min(2.8, Math.max(1.5, Math.abs(delta) / 80));
      this.element.dataset.liquidState = 'moving';
      this.motion.start(this);
    }
  }
  press() {
    if (!this.initialized || this.motion.preference.matches) return;
    this.pressure.target = .87;
    this.element.dataset.liquidState = 'moving';
    this.motion.start(this);
  }
  release() {
    if (!this.initialized) return;
    if (this.pressure.target === 1 && this.pressure.value === 1 && this.pressure.velocity === 0) return;
    this.pressure.target = 1;
    this.motion.start(this);
  }
  step(dt) {
    let moving = false;
    for (const spring of this.springs) moving = spring.step(dt) || moving;
    this.paint();
    if (!moving) this.element.dataset.liquidState = 'idle';
    return moving;
  }
  finish() {
    this.springs.forEach(spring => spring.snap());
    this.paint();
    this.element.dataset.liquidState = 'idle';
  }
  geometry() {
    const targetWidth = this.right.target - this.left.target;
    const speed = Math.abs((this.left.velocity + this.right.velocity) * .5);
    const pressure = clamp(this.pressure.value - speed / 14000, .78, 1.065);
    const width = clamp(this.right.value - this.left.value, targetWidth * .78, targetWidth + 70);
    const height = Math.max(1, this.height.value * pressure);
    let left = (this.left.value + this.right.value - width) * .5;
    if (this.clampEdges && this.container) left = clamp(left, 2, Math.max(2, this.container.clientWidth - width - 2));
    const top = this.top.value + (this.height.value - height) * .5;
    return { left, top, width, height, speed, pressure };
  }
  paint() {
    this.currentGeometry = this.geometry();
    const { left, top, width, height, speed, pressure } = this.currentGeometry;
    this.element.style.transform = `translate3d(${left.toFixed(3)}px,${top.toFixed(3)}px,0)`;
    this.element.style.width = `${width.toFixed(3)}px`;
    this.element.style.height = `${height.toFixed(3)}px`;
    this.element.style.borderRadius = `${Math.min(this.radius, height / 2).toFixed(3)}px`;
    if (this.container) {
      const strain = clamp(speed / 1200 + Math.abs(1 - pressure), 0, 1);
      this.container.style.scale = `${1 + strain * .008} ${1 - strain * .012}`;
    }
  }
}

// A knob expands under the finger, stretches in the direction of travel, and
// returns as one continuous lens. No separate CSS animation can drift from it.
export class LiquidKnob extends LiquidMarker {
  constructor(element, motion, options = {}) {
    super(element, motion, { radius: 40, clampEdges: false, ...options });
    this.engagement = new Spring(0, 440, 29);
    this.springs.push(this.engagement);
    this.dragging = false;
  }
  press() {
    if (!this.initialized || this.motion.preference.matches) return;
    this.engagement.target = 1;
    this.element.dataset.liquidState = 'moving';
    this.motion.start(this);
  }
  release() {
    if (!this.initialized) return;
    this.dragging = false;
    this.engagement.target = 0;
    super.release();
    if (this.engagement.value !== 0 || this.engagement.velocity !== 0) this.motion.start(this);
  }
  geometry() {
    const rect = super.geometry();
    const parent = this.element.parentElement;
    const parentStyle = getComputedStyle(parent);
    const parentWidth = parseFloat(parentStyle.width) || parent.clientWidth;
    const parentHeight = parseFloat(parentStyle.height) || parent.clientHeight;
    const size = this.right.target - this.left.target;
    const expansion = clamp(this.engagement.value, -.10, 1.08);
    const stretch = clamp(rect.speed / 650, 0, .18);
    const width = Math.min(parentWidth - 2, rect.width + size * (.30 * expansion + stretch));
    const height = Math.min(parentHeight - 2, rect.height * (1 + .09 * expansion));
    const left = clamp(rect.left + (rect.width - width) / 2, 1, Math.max(1, parentWidth - width - 1));
    return { ...rect, left, top: (parentHeight - height) / 2, width, height };
  }
}

export class LiquidToggle extends LiquidKnob {
  paint() {
    super.paint();
    const track = this.element.parentElement;
    const size = this.right.target - this.left.target;
    const inset = (track.clientHeight - size) / 2;
    const travel = Math.max(1, track.clientWidth - 2 * inset - size);
    const center = (this.left.value + this.right.value) / 2;
    track.style.setProperty('--switch-progress', clamp((center - inset - size / 2) / travel, 0, 1).toFixed(4));
  }
  bind(input) {
    let gesture = null;
    let suppressClick = false;
    const position = event => {
      const box = input.getBoundingClientRect();
      return clamp((event.clientX - box.left - box.height / 2) / Math.max(1, box.width - box.height), 0, 1);
    };
    input.addEventListener('pointerdown', event => {
      if (event.button !== 0) return;
      gesture = { id: event.pointerId, x: event.clientX, start: input.checked, moved: false };
      input.setPointerCapture(event.pointerId);
      this.press();
    });
    input.addEventListener('pointermove', event => {
      if (!gesture || gesture.id !== event.pointerId) return;
      gesture.moved ||= Math.abs(event.clientX - gesture.x) > 3;
      if (!gesture.moved) return;
      event.preventDefault();
      this.dragging = true;
      this.preview = position(event);
      const size = this.right.target - this.left.target;
      const track = this.element.parentElement;
      const inset = (track.clientHeight - size) / 2;
      const left = inset + (track.clientWidth - inset * 2 - size) * this.preview;
      this.left.snap(left);
      this.right.snap(left + size);
      this.paint();
      this.motion.animateFor(0);
    });
    const finish = (event, cancelled = false) => {
      if (!gesture || gesture.id !== event.pointerId) return;
      if (gesture.moved) {
        input.checked = cancelled ? gesture.start : this.preview >= .5;
        // Use the existing change handler so persistent state and optics agree.
        input.dispatchEvent(new Event('change', { bubbles: true }));
        suppressClick = !cancelled;
        if (suppressClick) setTimeout(() => { suppressClick = false; }, 0);
      }
      gesture = null;
      this.release();
    };
    input.addEventListener('pointerup', event => finish(event));
    input.addEventListener('pointercancel', event => finish(event, true));
    input.addEventListener('click', event => {
      if (!suppressClick) return;
      event.preventDefault();
      suppressClick = false;
    });
  }
}

export class LiquidRange extends LiquidKnob {
  constructor(element, motion, input, large) {
    super(element, motion);
    this.input = input;
    this.large = large;
  }
  update(immediate = false) {
    const parent = this.element.parentElement;
    const width = parent.checkVisibility() ? parseFloat(getComputedStyle(parent).width) || parent.clientWidth : 0;
    if (!width) { this.element.hidden = true; return; }
    const size = this.large ? 32 : 20;
    const percent = (Number(this.input.value) - Number(this.input.min)) / (Number(this.input.max) - Number(this.input.min));
    const x = (width - size - 4) * percent + 2;
    const rect = { x, y: 2, width: size, height: size };
    this.to(rect, immediate);
    if (this.dragging) {
      // Interpolate native value steps with a fast, damped follower. Snapping
      // erased all travel velocity and made the supposedly liquid knob rigid.
      this.left.stiffness = 1000; this.right.stiffness = 1100;
      this.left.damping = this.right.damping = 48;
    }
  }
  press() { this.dragging = true; super.press(); }
  paint() {
    super.paint();
    const track = this.element.parentElement;
    const center = this.currentGeometry.left + this.currentGeometry.width / 2;
    const width = parseFloat(getComputedStyle(track).width) || track.clientWidth;
    const fill = clamp(center / Math.max(1, width), 0, 1);
    track.style.setProperty('--visual-fill', `${fill * 100}%`);
    track.dataset.fill = fill.toFixed(5);
  }
}

export class LiquidSurface {
  constructor(element, motion) {
    this.element = element;
    this.motion = motion;
    this.scale = new Spring(1, 540, 32);
  }
  press() {
    if (this.motion.preference.matches) return;
    this.scale.target = .955;
    this.motion.start(this);
  }
  release() {
    if (this.scale.target === 1 && this.scale.value === 1) return;
    this.scale.target = 1;
    this.motion.start(this);
  }
  step(dt) { const active = this.scale.step(dt); this.paint(); return active; }
  paint() { this.element.style.scale = this.scale.value.toFixed(5); }
  finish() { this.scale.snap(1); this.paint(); }
}
