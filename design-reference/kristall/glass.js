// The photograph is rendered once, then sampled through rounded optical lenses.
// HTML above this canvas provides the accessible, interactive control layer.
const vertexSource = `
attribute vec2 a_position;
void main() { gl_Position = vec4(a_position, 0.0, 1.0); }
`;

const fragmentSource = `
precision highp float;
uniform sampler2D u_image;
uniform vec2 u_resolution;
uniform vec2 u_imageSize;
uniform vec4 u_rect;
uniform vec2 u_pixelRatio;
uniform float u_radius;
uniform float u_mode;
uniform float u_refraction;
uniform float u_frost;
uniform float u_tint;
uniform float u_brightness;
uniform float u_contrast;
uniform float u_focus;
uniform float u_readability;
uniform float u_control;
uniform float u_opacity;
uniform vec4 u_range;
uniform vec4 u_knob;
uniform float u_fill;
uniform vec2 u_pointer;

vec2 imageUV(vec2 p) {
  vec2 uv = p / u_resolution;
  float screenRatio = u_resolution.x / u_resolution.y;
  float imageRatio = u_imageSize.x / u_imageSize.y;
  if (screenRatio > imageRatio) uv.y = (uv.y - 0.5) * imageRatio / screenRatio + 0.5;
  else uv.x = (uv.x - 0.5) * screenRatio / imageRatio + 0.5;
  return clamp(uv, vec2(0.001), vec2(0.999));
}
vec3 scene(vec2 p) {
  vec3 color = texture2D(u_image, imageUV(p)).rgb;
  // A photographic exposure adjustment keeps highlights and reflections intact.
  color *= u_brightness;
  color = mix(color, color * vec3(0.70, 0.78, 0.77), u_focus * 0.6);
  return color;
}
float roundedDistance(vec2 p, vec2 halfSize, float r) {
  vec2 q = abs(p) - halfSize + r;
  return length(max(q, 0.0)) + min(max(q.x, q.y), 0.0) - r;
}
void main() {
  vec2 p = vec2(gl_FragCoord.x / u_pixelRatio.x, u_resolution.y - gl_FragCoord.y / u_pixelRatio.y);
  if (u_mode < 0.5) { gl_FragColor = vec4(scene(p), 1.0); return; }
  vec2 halfSize = u_rect.zw * 0.5;
  vec2 local = p - u_rect.xy - halfSize;
  float radius = min(u_radius, min(halfSize.x, halfSize.y));
  float sdf = roundedDistance(local, halfSize, radius);
  float coverage = 1.0 - smoothstep(-0.8, 0.5, sdf);
  if (coverage <= 0.0) discard;

  float depth = max(-sdf, 0.0);
  float bevelWidth = min(radius * 0.85, 24.0);
  float bevel = 1.0 - smoothstep(0.0, bevelWidth, depth);
  // Analytic SDF normal avoids four additional distance evaluations per pixel.
  vec2 q = abs(local) - halfSize + radius;
  vec2 outside = max(q, 0.0);
  vec2 normal = length(outside) > 0.001
    ? normalize(outside) * sign(local)
    : (q.x > q.y ? vec2(sign(local.x), 0.0) : vec2(0.0, sign(local.y)));
  float lens = sin(bevel * 1.570796);
  // Keep the mapping monotonic through the bevel. A stronger offset folds the
  // photograph back onto itself and looks like several stacked silhouettes.
  vec2 displacement = -normal * lens * u_refraction * bevelWidth * 0.34;
  // A gentle central magnification and stronger peripheral bending create depth.
  vec2 samplePoint = p + displacement - local * (0.012 * u_refraction);
  float filled = 0.0;
  float knob = 0.0;
  float knobDistance = 1000.0;
  if (u_fill >= 0.0) {
    vec2 knobHalf = u_knob.zw * 0.5;
    vec2 knobLocal = p - u_knob.xy - knobHalf;
    knobDistance = roundedDistance(knobLocal, knobHalf, min(knobHalf.x, knobHalf.y));
    knob = 1.0 - smoothstep(-4.0, 7.0, knobDistance);
    float end = min(u_range.x + u_range.z, u_knob.x + u_knob.z);
    vec2 fillHalf = vec2(max(1.0, end - u_range.x), u_range.w) * 0.5;
    float fillDistance = roundedDistance(p - u_range.xy - fillHalf, fillHalf, min(fillHalf.x, fillHalf.y));
    filled = 1.0 - smoothstep(-6.0, 6.0, fillDistance);
    // The knob is a smooth bulge in this same lens, not an opaque second
    // photograph with an abrupt boundary over the filled capsule.
    samplePoint -= knobLocal * (0.055 * u_refraction * knob);
  }
  float blur = u_frost * 15.0;
  vec3 col = scene(samplePoint);
  if (u_frost > 0.001) {
    col *= 0.24;
    col += scene(samplePoint + vec2(blur, 0.0)) * 0.12;
    col += scene(samplePoint - vec2(blur, 0.0)) * 0.12;
    col += scene(samplePoint + vec2(0.0, blur)) * 0.12;
    col += scene(samplePoint - vec2(0.0, blur)) * 0.12;
    col += scene(samplePoint + vec2(blur, blur) * 0.707) * 0.07;
    col += scene(samplePoint - vec2(blur, blur) * 0.707) * 0.07;
    col += scene(samplePoint + vec2(blur, -blur) * 0.707) * 0.07;
    col += scene(samplePoint + vec2(-blur, blur) * 0.707) * 0.07;
  }
  // Sub-pixel chromatic separation is confined to the curved glass edge.
  float dispersion = lens * u_refraction * 0.35;
  col.r = mix(col.r, scene(samplePoint + normal * dispersion).r, bevel * 0.65);
  col.b = mix(col.b, scene(samplePoint - normal * dispersion).b, bevel * 0.65);
  col = mix(col, vec3(0.84, 0.92, 0.89), u_tint * 0.31 + u_frost * 0.10);
  col *= 0.91;
  col = mix(col, vec3(0.94, 0.98, 0.96), u_control * 0.14);
  if (u_fill >= 0.0) {
    col = mix(col, vec3(0.96, 0.99, 0.98), filled * 0.23 + knob * 0.13);
    col += exp(-pow(knobDistance / 4.2, 2.0)) * 0.085;
  }
  col = mix(col, vec3(0.055, 0.15, 0.13), u_contrast * 0.70);
  // Large popovers gain optical density in their interior. The transparent
  // refractive bevel stays clear, while labels retain contrast over bright sky.
  col = mix(col, vec3(0.027, 0.094, 0.078), u_readability * 0.66 * smoothstep(2.0, 18.0, depth));

  vec2 light = normalize(vec2(-0.7, -1.0) + (u_pointer - p) / max(u_resolution, vec2(1.0)) * 1.7);
  float highlight = pow(max(dot(normal, light), 0.0), 4.0);
  float backlight = pow(max(dot(normal, -light), 0.0), 5.0);
  float rim = exp(-depth * 0.8);
  col += vec3(0.95, 1.0, 0.98) * rim * (0.10 + highlight * 0.52 + backlight * 0.12);
  // A wide specular response moves with the pointer without obscuring content.
  vec2 spot = (p - u_pointer) / vec2(300.0, 230.0);
  col += exp(-dot(spot, spot) * 2.0) * 0.045;
  gl_FragColor = vec4(col, coverage * u_opacity);
}
`;

export class GlassRenderer {
  constructor(canvas) {
    this.canvas = canvas;
    this.gl = canvas.getContext('webgl', { alpha: false, antialias: false, depth: false, powerPreference: 'low-power', preserveDrawingBuffer: true });
    if (!this.gl) throw new Error('WebGL unavailable');
    this.settings = { refraction: 1, frost: 0, tint: .08, brightness: 1, contrast: 0, focus: 0, motion: true };
    this.pointer = [window.innerWidth * .2, 0];
    this.imageSize = [1, 1];
    this.pending = false;
    this.frameId = 0;
    this.sceneGeneration = 0;
    this.ready = false;
    this.init();
    this.resize = this.resize.bind(this);
    window.addEventListener('resize', this.resize);
    window.addEventListener('scroll', () => this.requestDraw(), { passive: true, capture: true });
    window.addEventListener('pointermove', event => {
      if (!this.settings.motion || matchMedia('(prefers-reduced-motion: reduce)').matches) return;
      this.pointer = [event.clientX, event.clientY];
      this.requestDraw();
    }, { passive: true });
    this.observer = new ResizeObserver(() => { this.resize(); this.requestDraw(); });
    this.observer.observe(canvas);
    document.querySelectorAll('[data-glass]:not(.liquid-marker)').forEach(el => this.observer.observe(el));
    canvas.addEventListener('webglcontextlost', event => {
      event.preventDefault();
      this.ready = false;
      cancelAnimationFrame(this.frameId);
      this.pending = false;
      document.documentElement.dataset.renderer = 'css';
      canvas.style.visibility = 'hidden';
      document.dispatchEvent(new CustomEvent('glass-fallback'));
    });
    canvas.addEventListener('webglcontextrestored', async () => {
      try { this.init(); await this.loadScene(this.currentScene || 'alpine'); }
      catch { document.documentElement.dataset.renderer = 'css'; }
    });
    this.resize();
  }

  init() {
    const gl = this.gl;
    const compile = (type, source) => {
      const shader = gl.createShader(type);
      gl.shaderSource(shader, source);
      gl.compileShader(shader);
      if (!gl.getShaderParameter(shader, gl.COMPILE_STATUS)) throw new Error(gl.getShaderInfoLog(shader));
      return shader;
    };
    this.program = gl.createProgram();
    const vertex = compile(gl.VERTEX_SHADER, vertexSource);
    const fragment = compile(gl.FRAGMENT_SHADER, fragmentSource);
    gl.attachShader(this.program, vertex);
    gl.attachShader(this.program, fragment);
    gl.linkProgram(this.program);
    gl.deleteShader(vertex);
    gl.deleteShader(fragment);
    if (!gl.getProgramParameter(this.program, gl.LINK_STATUS)) throw new Error(gl.getProgramInfoLog(this.program));
    gl.useProgram(this.program);
    const buffer = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, buffer);
    gl.bufferData(gl.ARRAY_BUFFER, new Float32Array([-1,-1, 3,-1, -1,3]), gl.STATIC_DRAW);
    const position = gl.getAttribLocation(this.program, 'a_position');
    gl.enableVertexAttribArray(position);
    gl.vertexAttribPointer(position, 2, gl.FLOAT, false, 0, 0);
    this.uniforms = {};
    ['image','resolution','imageSize','rect','pixelRatio','radius','mode','refraction','frost','tint','brightness','contrast','focus','readability','control','opacity','range','knob','fill','pointer'].forEach(name => {
      this.uniforms[name] = gl.getUniformLocation(this.program, `u_${name}`);
    });
    gl.activeTexture(gl.TEXTURE0);
    this.texture = gl.createTexture();
    gl.bindTexture(gl.TEXTURE_2D, this.texture);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.CLAMP_TO_EDGE);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.CLAMP_TO_EDGE);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.LINEAR);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.LINEAR);
    gl.uniform1i(this.uniforms.image, 0);
  }

  async loadScene(name) {
    const generation = ++this.sceneGeneration;
    const image = new Image();
    image.src = new URL(`assets/${name}.jpg`, import.meta.url).href;
    await image.decode();
    if (generation !== this.sceneGeneration || this.gl.isContextLost()) return false;
    const gl = this.gl;
    gl.bindTexture(gl.TEXTURE_2D, this.texture);
    gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGB, gl.RGB, gl.UNSIGNED_BYTE, image);
    this.imageSize = [image.naturalWidth, image.naturalHeight];
    this.currentScene = name;
    this.ready = true;
    document.documentElement.dataset.renderer = 'webgl';
    this.canvas.style.visibility = 'visible';
    document.dispatchEvent(new CustomEvent('glass-ready'));
    this.requestDraw();
    return true;
  }

  resize() {
    this.dpr = Math.min(window.devicePixelRatio || 1, 1.5);
    // innerWidth includes a classic scrollbar; the fixed CSS canvas does not.
    // Measure the actual drawing surface so lenses and HTML share coordinates.
    const bounds = this.canvas.getBoundingClientRect();
    this.width = bounds.width;
    this.height = bounds.height;
    const width = Math.round(this.width * this.dpr);
    const height = Math.round(this.height * this.dpr);
    if (this.canvas.width !== width) this.canvas.width = width;
    if (this.canvas.height !== height) this.canvas.height = height;
    this.pixelRatio = [width / this.width, height / this.height];
    this.gl.viewport(0, 0, this.canvas.width, this.canvas.height);
    this.requestDraw();
  }

  update(settings) {
    Object.assign(this.settings, settings);
    this.requestDraw();
  }

  requestDraw() {
    if (this.pending || !this.ready) return;
    this.pending = true;
    this.frameId = requestAnimationFrame(() => {
      this.pending = false;
      this.draw();
    });
  }

  draw() {
    if (!this.ready || document.hidden || this.gl.isContextLost()) return;
    const gl = this.gl;
    const u = this.uniforms;
    gl.useProgram(this.program);
    gl.uniform2f(u.resolution, this.width, this.height);
    gl.uniform2f(u.imageSize, ...this.imageSize);
    gl.uniform2f(u.pointer, ...this.pointer);
    gl.uniform2f(u.pixelRatio, ...this.pixelRatio);
    for (const key of ['refraction','frost','tint','brightness','contrast','focus']) gl.uniform1f(u[key], this.settings[key]);
    gl.disable(gl.SCISSOR_TEST);
    gl.disable(gl.BLEND);
    gl.uniform1f(u.mode, 0);
    gl.drawArrays(gl.TRIANGLES, 0, 3);
    gl.uniform1f(u.mode, 1);
    gl.enable(gl.SCISSOR_TEST);
    gl.enable(gl.BLEND);
    gl.blendFunc(gl.SRC_ALPHA, gl.ONE_MINUS_SRC_ALPHA);
    // DOM order is also the painter's order; popovers follow widgets in the DOM.
    document.querySelectorAll('[data-glass]').forEach(element => {
      if (element.classList.contains('range-thumb') && element.closest('.filled-range')) return;
      if (!element.checkVisibility({ checkVisibilityCSS: true, checkOpacity: true })) return;
      const rect = element.getBoundingClientRect();
      if (rect.bottom < 0 || rect.top > this.height || rect.width === 0) return;
      const style = getComputedStyle(element);
      const scale = Math.min(rect.width / Math.max(1, parseFloat(style.width) || element.offsetWidth), rect.height / Math.max(1, parseFloat(style.height) || element.offsetHeight));
      const radius = (parseFloat(style.borderTopLeftRadius) || 24) * scale;
      gl.uniform4f(u.rect, rect.x, rect.y, rect.width, rect.height);
      const popover = element.closest('.floating-panel');
      gl.uniform1f(u.readability, popover ? 1 : 0);
      gl.uniform1f(u.control, element.classList.contains('liquid-marker') || element.classList.contains('range-thumb') ? 1 : 0);
      let opacity = Number(style.opacity);
      if (popover && popover !== element) opacity *= Number(getComputedStyle(popover).opacity);
      const workspace = element.closest('.workspace');
      if (workspace && workspace !== element) opacity *= Number(getComputedStyle(workspace).opacity);
      const intro = element.closest('.intro');
      if (intro && intro !== element) opacity *= Number(getComputedStyle(intro).opacity);
      gl.uniform1f(u.opacity, opacity);
      const range = element.closest('.filled-range');
      if (range) {
        const rangeRect = range.getBoundingClientRect();
        const knobRect = range.querySelector('.range-thumb').getBoundingClientRect();
        gl.uniform4f(u.range, rangeRect.x, rangeRect.y, rangeRect.width, rangeRect.height);
        gl.uniform4f(u.knob, knobRect.x, knobRect.y, knobRect.width, knobRect.height);
        gl.uniform1f(u.fill, Number(range.dataset.fill || 0));
      } else gl.uniform1f(u.fill, -1);
      gl.uniform1f(u.radius, element.dataset.glass === 'circle' ? rect.width / 2 : radius);
      let clip = { left: 0, right: this.width, top: 0, bottom: this.height };
      // The mobile content viewport reserves a separate area for the dock.
      // Scissor the canvas lenses exactly where CSS clips their HTML controls.
      if (matchMedia('(max-width: 760px)').matches && element.closest('.home-view')) clip = document.querySelector('main').getBoundingClientRect();
      const [pixelX, pixelY] = this.pixelRatio;
      const left = Math.max(0, Math.floor(Math.max(rect.left, clip.left) * pixelX) - 1);
      const bottom = Math.max(0, Math.floor((this.height - Math.min(rect.bottom, clip.bottom)) * pixelY));
      const right = Math.min(this.canvas.width, Math.ceil(Math.min(rect.right, clip.right) * pixelX) + 1);
      const top = Math.min(this.canvas.height, Math.ceil((this.height - Math.max(rect.top, clip.top)) * pixelY));
      gl.scissor(left, bottom, Math.max(0, right - left), Math.max(0, top - bottom));
      gl.drawArrays(gl.TRIANGLES, 0, 3);
    });
    gl.disable(gl.SCISSOR_TEST);
  }
}
