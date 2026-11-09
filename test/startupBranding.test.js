const fs = require('fs');
const path = require('path');
const test = require('node:test');
const assert = require('node:assert/strict');

const projectRoot = path.join(__dirname, '..');

function decodeRgba(png) {
	assert.equal(png[24],8);assert.equal(png[25],6,'RGBA PNG with a real alpha channel');
	const width=png.readUInt32BE(16),height=png.readUInt32BE(20),chunks=[];
	for(let p=8;p<png.length;) {const length=png.readUInt32BE(p);if(png.toString('ascii',p+4,p+8)==='IDAT')chunks.push(png.subarray(p+8,p+8+length));p+=length+12;}
	const packed=require('node:zlib').inflateSync(Buffer.concat(chunks)),stride=width*4,pixels=Buffer.alloc(height*stride);
	const paeth=(a,b,c)=>{const p=a+b-c,pa=Math.abs(p-a),pb=Math.abs(p-b),pc=Math.abs(p-c);return pa<=pb&&pa<=pc ? a : pb<=pc ? b : c;};
	for(let y=0;y<height;y++)for(let x=0;x<stride;x++) {
		const filter=packed[y*(stride+1)],left=x>=4 ? pixels[y*stride+x-4] : 0,up=y ? pixels[(y-1)*stride+x] : 0,corner=y&&x>=4 ? pixels[(y-1)*stride+x-4] : 0;
		const predict=[0,left,up,Math.floor((left+up)/2),paeth(left,up,corner)][filter];assert.notEqual(predict,undefined);
		pixels[y*stride+x]=(packed[y*(stride+1)+x+1]+predict)&255;
	}
	return {width,height,alpha:(x,y)=>pixels[(y*width+x)*4+3]};
}

test('produktives Intro verwendet die Ein-Text-Animation aus Variante 1', () => {
	const html = fs.readFileSync(path.join(projectRoot, 'public', 'splash.html'), 'utf8');
	const script = fs.readFileSync(path.join(projectRoot, 'public', 'splash.js'), 'utf8');
	const css = fs.readFileSync(path.join(projectRoot, 'public', 'style.css'), 'utf8');

	assert.equal((html.match(/>Repetierer<\/text>/g) || []).length, 1);
	assert.match(html, /class="startup-loader-word"[\s\S]+mask="url\(#startup-line-reveal-mask\)"/);
	assert.match(html, /class="startup-fill-motion"[^>]+begin="indefinite"[^>]+dur="1\.2s"/);
	assert.match(html, /class="startup-line-reveal-motion"[^>]+begin="indefinite"[^>]+dur="3\.5s"/);
	assert.match(script, /lineStart: 140/);
	assert.match(script, /lineComplete: 3640/);
	assert.match(script, /fillComplete: 4840/);
	assert.match(script, /logoComplete: 6390/);
	assert.match(css, /\.startup-line-drawing \.startup-loader-word[\s\S]+startup-line-draw 3\.5s linear forwards/);
	assert.match(css, /\.startup-word-filled \.startup-loader-word[\s\S]+translateX\(-40px\)/);
	assert.match(css, /\.startup-seven-visible \.startup-loader-seven[\s\S]+startup-seven-arrive/);
	assert.equal((css.match(/format\("truetype"\)/g) || []).length, 6);
	assert.doesNotMatch(css, /format\("ttf"\)/);
});

test('Windows-App und Update-Shortcuts verwenden das neue Icon stabil', () => {
	const packageJson = JSON.parse(fs.readFileSync(path.join(projectRoot, 'package.json'), 'utf8'));
	const main = fs.readFileSync(path.join(projectRoot, 'app.js'), 'utf8');

	assert.equal(packageJson.build.win.icon, './public/icons/icon.ico');
	assert.equal(packageJson.build.nsis.installerIcon, './public/icons/icon.ico');
	assert.equal(packageJson.build.nsis.uninstallerIcon, './public/icons/icon.ico');
	assert.equal(packageJson.build.nsis.createDesktopShortcut, 'always');
	assert.match(main, /app\.setAppUserModelId\('net\.srpnt3\.repetierer'\)/);
	assert.equal((main.match(/public', 'icons', 'icon\.ico'/g) || []).length, 2);
});

test('new R arrow icon retains the full PNG and ten 32-bit ICO sizes through 256px', () => {
	const png = fs.readFileSync(path.join(projectRoot,'public/icons/icon.png'));
	assert.equal(png.readUInt32BE(16),1254); assert.equal(png.readUInt32BE(20),1254);
	const image=decodeRgba(png);
	for(const [x,y] of [[0,0],[1253,0],[0,1253],[1253,1253]])assert.equal(image.alpha(x,y),0,'No black opaque corner in the app PNG');
	assert.ok(image.alpha(627,627)>240,'The logo interior remains visible');
	const ico = fs.readFileSync(path.join(projectRoot,'public/icons/icon.ico'));
	assert.equal(ico.readUInt16LE(0),0); assert.equal(ico.readUInt16LE(2),1); assert.equal(ico.readUInt16LE(4),10);
	const sizes = [];
	for (let i=0;i<10;i++) {
		const offset = 6+i*16, width = ico[offset] || 256, height = ico[offset+1] || 256;
		assert.equal(width,height); assert.equal(ico.readUInt16LE(offset+6),32);
		const payload = ico.subarray(ico.readUInt32LE(offset+12),ico.readUInt32LE(offset+12)+ico.readUInt32LE(offset+8));
		assert.equal(payload.subarray(1,4).toString(),'PNG');
		assert.equal(payload.readUInt32BE(16),width); assert.equal(payload.readUInt32BE(20),height);
		const frame=decodeRgba(payload);
		for(const [x,y] of [[0,0],[width-1,0],[0,height-1],[width-1,height-1]])assert.equal(frame.alpha(x,y),0,'Transparent ICO corner at '+width);
		sizes.push(width);
	}
	assert.deepEqual(sizes,[16,20,24,32,40,48,64,96,128,256]);
});
