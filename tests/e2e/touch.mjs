/** Real touch input through the DevTools protocol (Playwright's touchscreen only taps). */
export async function touch(page) {
  const cdp = await page.context().newCDPSession(page);
  const send = (type, pts) => cdp.send('Input.dispatchTouchEvent', { type, touchPoints: pts.map((p, i) => ({ x: p.x, y: p.y, id: i, radiusX: 4, radiusY: 4, force: 1 })) });
  const wait = (ms) => page.waitForTimeout(ms);
  return {
    cdp,
    async tap(x, y) {
      await send('touchStart', [{ x, y }]);
      await wait(40);
      await send('touchEnd', []);
      await wait(60);
    },
    async doubleTap(x, y) {
      await send('touchStart', [{ x, y }]);
      await send('touchEnd', []);
      await wait(50);
      await send('touchStart', [{ x, y }]);
      await send('touchEnd', []);
      await wait(80);
    },
    async longPress(x, y, ms = 700) {
      await send('touchStart', [{ x, y }]);
      await wait(ms);
      await send('touchEnd', []);
      await wait(60);
    },
    /** press, drag, release; `during` runs half-way with the finger still down */
    async drag(a, b, { steps = 12, stepMs = 16, hold = 0, settle = 0, during } = {}) {
      await send('touchStart', [a]);
      if (hold) await wait(hold);
      for (let i = 1; i <= steps; i++) {
        await send('touchMove', [{ x: a.x + ((b.x - a.x) * i) / steps, y: a.y + ((b.y - a.y) * i) / steps }]);
        await wait(stepMs);
        if (during && i === Math.ceil(steps / 2)) await during();
      }
      if (settle) await wait(settle); // finger rests before lifting: no momentum
      await send('touchEnd', []);
      await wait(60);
    },
    /** quick flick that should start momentum scrolling */
    async fling(a, b, { steps = 6, stepMs = 8 } = {}) {
      return this.drag(a, b, { steps, stepMs });
    },
    async pinch(c, from, to, { steps = 16, stepMs = 16 } = {}) {
      const pts = (d) => [{ x: c.x - d / 2, y: c.y }, { x: c.x + d / 2, y: c.y }];
      await send('touchStart', pts(from));
      for (let i = 1; i <= steps; i++) {
        await send('touchMove', pts(from + ((to - from) * i) / steps));
        await wait(stepMs);
      }
      await send('touchEnd', []);
      await wait(100);
    },
    start: (pts) => send('touchStart', pts),
    move: (pts) => send('touchMove', pts),
    end: () => send('touchEnd', []),
  };
}
