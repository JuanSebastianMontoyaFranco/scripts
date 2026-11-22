(function () {
  var t,
    e,
    n,
    i,
    o,
    r,
    s,
    a,
    u,
    l,
    d,
    h,
    p,
    c,
    m,
    f,
    g,
    v,
    w = [].slice;
  (t = /^\(?([^)]*)\)?(?:(.)(d+))?$/),
    (d = document.createElement("div").style),
    (i =
      null != d.transition ||
      null != d.webkitTransition ||
      null != d.mozTransition ||
      null != d.oTransition),
    (u =
      window.requestAnimationFrame ||
      window.mozRequestAnimationFrame ||
      window.webkitRequestAnimationFrame ||
      window.msRequestAnimationFrame),
    (e =
      window.MutationObserver ||
      window.WebKitMutationObserver ||
      window.MozMutationObserver),
    (r = function (t) {
      var e;
      return ((e = document.createElement("div")).innerHTML = t), e.children[0];
    }),
    (a = function (t, e) {
      return (t.className = t.className.replace(
        new RegExp("(^| )" + e.split(" ").join("|") + "( |$)", "gi"),
        " "
      ));
    }),
    (o = function (t, e) {
      return a(t, e), (t.className += " " + e);
    }),
    (h = function (t, e) {
      var n;
      if (null != document.createEvent)
        return (
          (n = document.createEvent("HTMLEvents")).initEvent(e, !0, !0),
          t.dispatchEvent(n)
        );
    }),
    (s = function () {
      var t, e;
      return null !=
        (t =
          null != (e = window.performance) && "function" == typeof e.now
            ? e.now()
            : void 0)
        ? t
        : +new Date();
    }),
    (l = function (t, e) {
      return (
        null == e && (e = 0),
        e
          ? ((t *= Math.pow(10, e)),
            (t += 0.5),
            (t = Math.floor(t)) / Math.pow(10, e))
          : Math.round(t)
      );
    }),
    (p = function (t) {
      return t < 0 ? Math.ceil(t) : Math.floor(t);
    }),
    (m = !1),
    (c = function () {
      var t, e, n, i, o;
      if (!m && null != window.jQuery) {
        for (
          m = !0, o = [], e = 0, n = (i = ["html", "text"]).length;
          e < n;
          e++
        )
          (t = i[e]),
            o.push(
              (function (t) {
                var e;
                return (
                  (e = window.jQuery.fn[t]),
                  (window.jQuery.fn[t] = function (t) {
                    var n;
                    return null == t ||
                      null == (null != (n = this[0]) ? n.odometer : void 0)
                      ? e.apply(this, arguments)
                      : this[0].odometer.update(t);
                  })
                );
              })(t)
            );
        return o;
      }
    })(),
    setTimeout(c, 0),
    (n = (function () {
      function n(t) {
        var e,
          i,
          o,
          r,
          s,
          a,
          u,
          l,
          d,
          h = this;
        if (
          ((this.options = t),
          (this.el = this.options.el),
          null != this.el.odometer)
        )
          return this.el.odometer;
        for (e in ((this.el.odometer = this), (u = n.options)))
          (o = u[e]), null == this.options[e] && (this.options[e] = o);
        null == (r = this.options).duration && (r.duration = 2e3),
          (this.MAX_VALUES =
            (this.options.duration / 33.333333333333336 / 2) | 0),
          this.resetFormat(),
          (this.value = this.cleanValue(
            null != (l = this.options.value) ? l : ""
          )),
          this.renderInside(),
          this.render();
        try {
          for (
            s = 0, a = (d = ["innerHTML", "innerText", "textContent"]).length;
            s < a;
            s++
          )
            (i = d[s]),
              null != this.el[i] &&
                (function (t) {
                  Object.defineProperty(h.el, t, {
                    get: function () {
                      var e;
                      return "innerHTML" === t
                        ? h.inside.outerHTML
                        : null != (e = h.inside.innerText)
                        ? e
                        : h.inside.textContent;
                    },
                    set: function (t) {
                      return h.update(t);
                    },
                  });
                })(i);
        } catch (t) {
          t, this.watchForMutations();
        }
      }
      return (
        (n.prototype.renderInside = function () {
          return (
            (this.inside = document.createElement("div")),
            (this.inside.className = "odometer-inside"),
            (this.el.innerHTML = ""),
            this.el.appendChild(this.inside)
          );
        }),
        (n.prototype.watchForMutations = function () {
          var t = this;
          if (null != e)
            try {
              return (
                null == this.observer &&
                  (this.observer = new e(function (e) {
                    var n;
                    return (
                      (n = t.el.innerText),
                      t.renderInside(),
                      t.render(t.value),
                      t.update(n)
                    );
                  })),
                (this.watchMutations = !0),
                this.startWatchingMutations()
              );
            } catch (t) {
              t;
            }
        }),
        (n.prototype.startWatchingMutations = function () {
          if (this.watchMutations)
            return this.observer.observe(this.el, { childList: !0 });
        }),
        (n.prototype.stopWatchingMutations = function () {
          var t;
          return null != (t = this.observer) ? t.disconnect() : void 0;
        }),
        (n.prototype.cleanValue = function (t) {
          var e;
          return (
            "string" == typeof t &&
              ((t = (t = (t = t.replace(
                null != (e = this.format.radix) ? e : ".",
                "<radix>"
              )).replace(/[.,]/g, "")).replace("<radix>", ".")),
              (t = parseFloat(t, 10) || 0)),
            l(t, this.format.precision)
          );
        }),
        (n.prototype.bindTransitionEnd = function () {
          var t,
            e,
            n,
            i,
            o,
            r,
            s = this;
          if (!this.transitionEndBound) {
            for (
              this.transitionEndBound = !0,
                e = !1,
                r = [],
                n = 0,
                i = (o =
                  "transitionend webkitTransitionEnd oTransitionEnd otransitionend MSTransitionEnd".split(
                    " "
                  )).length;
              n < i;
              n++
            )
              (t = o[n]),
                r.push(
                  this.el.addEventListener(
                    t,
                    function () {
                      return (
                        e ||
                          ((e = !0),
                          setTimeout(function () {
                            return (
                              s.render(), (e = !1), h(s.el, "odometerdone")
                            );
                          }, 0)),
                        !0
                      );
                    },
                    !1
                  )
                );
            return r;
          }
        }),
        (n.prototype.resetFormat = function () {
          var e, n, i, o, r, s, a, u;
          if (
            ((e = null != (a = this.options.format) ? a : "(,ddd).dd") ||
              (e = "d"),
            !(i = t.exec(e)))
          )
            throw new Error("Odometer: Unparsable digit format");
          return (
            (s = (u = i.slice(1, 4))[0]),
            (r = u[1]),
            (o = (null != (n = u[2]) ? n.length : void 0) || 0),
            (this.format = { repeating: s, radix: r, precision: o })
          );
        }),
        (n.prototype.render = function (t) {
          var e, n, o, r, s, a, u, l, d, h, p, c;
          for (
            null == t && (t = this.value),
              this.stopWatchingMutations(),
              this.resetFormat(),
              this.inside.innerHTML = "",
              a = this.options.theme,
              s = [],
              l = 0,
              h = (e = this.el.className.split(" ")).length;
            l < h;
            l++
          )
            (n = e[l]).length &&
              ((r = /^odometer-theme-(.+)$/.exec(n))
                ? (a = r[1])
                : /^odometer(-|$)/.test(n) || s.push(n));
          for (
            s.push("odometer"),
              i || s.push("odometer-no-transitions"),
              a ? s.push("odometer-theme-" + a) : s.push("odometer-auto-theme"),
              this.el.className = s.join(" "),
              this.ribbons = {},
              this.digits = [],
              u = !this.format.precision,
              d = 0,
              p = (c = t.toFixed(this.format.precision).split("").reverse())
                .length;
            d < p;
            d++
          )
            "." === (o = c[d]) && (u = !0), this.addDigit(o, u);
          return this.startWatchingMutations();
        }),
        (n.prototype.update = function (t) {
          var e,
            n = this;
          if ((e = (t = this.cleanValue(t)) - this.value))
            return (
              a(
                this.el,
                "odometer-animating-up odometer-animating-down odometer-animating"
              ),
              o(
                this.el,
                e > 0 ? "odometer-animating-up" : "odometer-animating-down"
              ),
              this.stopWatchingMutations(),
              this.animate(t),
              this.startWatchingMutations(),
              setTimeout(function () {
                return n.el.offsetHeight, o(n.el, "odometer-animating");
              }, 0),
              (this.value = t)
            );
        }),
        (n.prototype.renderDigit = function () {
          return r(
            '<span class="odometer-digit"><span class="odometer-digit-spacer">0</span><span class="odometer-digit-inner"><span class="odometer-ribbon"><span class="odometer-ribbon-inner"><span class="odometer-value"></span></span></span></span></span>'
          );
        }),
        (n.prototype.insertDigit = function (t, e) {
          return null != e
            ? this.inside.insertBefore(t, e)
            : this.inside.children.length
            ? this.inside.insertBefore(t, this.inside.children[0])
            : this.inside.appendChild(t);
        }),
        (n.prototype.addSpacer = function (t, e, n) {
          var i;
          return (
            ((i = r(
              '<span class="odometer-formatting-mark"></span>'
            )).innerHTML = t),
            n && o(i, n),
            this.insertDigit(i, e)
          );
        }),
        (n.prototype.addDigit = function (t, e) {
          var n, i, o, r;
          if ((null == e && (e = !0), "-" === t))
            return this.addSpacer(t, null, "odometer-negation-mark");
          if ("." === t)
            return this.addSpacer(
              null != (r = this.format.radix) ? r : ".",
              null,
              "odometer-radix-mark"
            );
          if (e)
            for (o = !1; ; ) {
              if (!this.format.repeating.length) {
                if (o) throw new Error("Bad odometer format without digits");
                this.resetFormat(), (o = !0);
              }
              if (
                ((n = this.format.repeating[this.format.repeating.length - 1]),
                (this.format.repeating = this.format.repeating.substring(
                  0,
                  this.format.repeating.length - 1
                )),
                "d" === n)
              )
                break;
              this.addSpacer(n);
            }
          return (
            ((i = this.renderDigit()).querySelector(
              ".odometer-value"
            ).innerHTML = t),
            this.digits.push(i),
            this.insertDigit(i)
          );
        }),
        (n.prototype.animate = function (t) {
          return i && "count" !== this.options.animation
            ? this.animateSlide(t)
            : this.animateCount(t);
        }),
        (n.prototype.animateCount = function (t) {
          var e,
            n,
            i,
            o,
            r,
            a = this;
          if ((n = +t - this.value))
            return (
              (o = i = s()),
              (e = this.value),
              (r = function () {
                var l, d;
                return s() - o > a.options.duration
                  ? ((a.value = t), a.render(), void h(a.el, "odometerdone"))
                  : ((l = s() - i) > 50 &&
                      ((i = s()),
                      (d = l / a.options.duration),
                      (e += n * d),
                      a.render(Math.round(e))),
                    null != u ? u(r) : setTimeout(r, 50));
              })()
            );
        }),
        (n.prototype.getDigitCount = function () {
          var t, e, n, i, o, r;
          for (
            t = o = 0,
              r = (i = 1 <= arguments.length ? w.call(arguments, 0) : [])
                .length;
            o < r;
            t = ++o
          )
            (n = i[t]), (i[t] = Math.abs(n));
          return (
            (e = Math.max.apply(Math, i)),
            Math.ceil(Math.log(e + 1) / Math.log(10))
          );
        }),
        (n.prototype.getFractionalDigitCount = function () {
          var t, e, n, i, o, r, s;
          for (
            e = /^\-?\d*\.(\d*?)$/,
              t = r = 0,
              s = (o = 1 <= arguments.length ? w.call(arguments, 0) : [])
                .length;
            r < s;
            t = ++r
          )
            (i = o[t]),
              (o[t] = i.toFixed(this.format.precision)),
              (n = e.exec(o[t])),
              (o[t] = null == n ? 0 : n[1].length);
          return Math.max.apply(Math, o);
        }),
        (n.prototype.resetDigits = function () {
          return (
            (this.digits = []),
            (this.ribbons = []),
            (this.inside.innerHTML = ""),
            this.resetFormat()
          );
        }),
        (n.prototype.animateSlide = function (t) {
          var e,
            n,
            i,
            r,
            s,
            a,
            u,
            l,
            d,
            h,
            c,
            m,
            f,
            g,
            v,
            w,
            M,
            y,
            b,
            T,
            E,
            x,
            L,
            S,
            D,
            F,
            A;
          (w = this.value),
            this.getFractionalDigitCount(w, t),
            (l = this.getFractionalDigitCount(w, t));
          var C = (p(t) + "").length,
            H = (p(w) + "").length,
            O = Math.max(C, H);
          if (
            (l && ((t *= Math.pow(10, l)), (w *= Math.pow(10, l))), (i = t - w))
          ) {
            for (
              this.bindTransitionEnd(), r = O + l, s = [], e = 0, c = b = 0;
              0 <= r ? b < r : b > r;
              c = 0 <= r ? ++b : --b
            ) {
              if (
                ((M = p(w / Math.pow(10, r - c - 1))),
                (a = (u = p(t / Math.pow(10, r - c - 1))) - M),
                Math.abs(a) > this.MAX_VALUES)
              ) {
                for (
                  h = [],
                    m = a / (this.MAX_VALUES + this.MAX_VALUES * e * 0.5),
                    n = M;
                  (a > 0 && n < u) || (a < 0 && n > u);

                )
                  h.push(Math.round(n)), (n += m);
                h[h.length - 1] !== u && h.push(u), e++;
              } else
                h = function () {
                  A = [];
                  for (var t = M; M <= u ? t <= u : t >= u; M <= u ? t++ : t--)
                    A.push(t);
                  return A;
                }.apply(this);
              for (c = T = 0, x = h.length; T < x; c = ++T)
                (d = h[c]), (h[c] = Math.abs(d % 10));
              s.push(h);
            }
            for (
              this.resetDigits(), c = E = 0, L = (F = s.reverse()).length;
              E < L;
              c = ++E
            )
              for (
                h = F[c],
                  this.digits[c] || this.addDigit(" ", c >= l),
                  null == (y = this.ribbons)[c] &&
                    (y[c] = this.digits[c].querySelector(
                      ".odometer-ribbon-inner"
                    )),
                  this.ribbons[c].innerHTML = "",
                  i < 0 && (h = h.reverse()),
                  f = D = 0,
                  S = h.length;
                D < S;
                f = ++D
              )
                (d = h[f]),
                  ((v = document.createElement("div")).className =
                    "odometer-value"),
                  (v.innerHTML = d),
                  this.ribbons[c].appendChild(v),
                  f === h.length - 1 && o(v, "odometer-last-value"),
                  0 === f && o(v, "odometer-first-value");
            return (
              M < 0 && this.addDigit("-"),
              null != (g = this.inside.querySelector(".odometer-radix-mark")) &&
                g.parent.removeChild(g),
              l
                ? this.addSpacer(
                    this.format.radix,
                    this.digits[l - 1],
                    "odometer-radix-mark"
                  )
                : void 0
            );
          }
        }),
        n
      );
    })()),
    (n.options = null != (g = window.odometerOptions) ? g : {}),
    setTimeout(function () {
      var t, e, i, o, r;
      if (window.odometerOptions) {
        for (t in ((r = []), (o = window.odometerOptions)))
          (e = o[t]),
            r.push(
              null != (i = n.options)[t] ? (i = n.options)[t] : (i[t] = e)
            );
        return r;
      }
    }, 0),
    (n.init = function () {
      var t, e, i, o, r, s;
      if (null != document.querySelectorAll) {
        for (
          s = [],
            i = 0,
            o = (e = document.querySelectorAll(
              n.options.selector || ".odometer"
            )).length;
          i < o;
          i++
        )
          (t = e[i]),
            s.push(
              (t.odometer = new n({
                el: t,
                value: null != (r = t.innerText) ? r : t.textContent,
              }))
            );
        return s;
      }
    }),
    null != (null != (v = document.documentElement) ? v.doScroll : void 0) &&
    null != document.createEventObject
      ? ((f = document.onreadystatechange),
        (document.onreadystatechange = function () {
          return (
            "complete" === document.readyState &&
              !1 !== n.options.auto &&
              n.init(),
            null != f ? f.apply(this, arguments) : void 0
          );
        }))
      : document.addEventListener(
          "DOMContentLoaded",
          function () {
            if (!1 !== n.options.auto) return n.init();
          },
          !1
        ),
    (window.Odometer = n);
}.call(this));
