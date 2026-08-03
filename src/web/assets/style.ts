export const WEB_APP_CSS = String.raw`
:root {
  color-scheme: light;
  --canvas: #f5f6f8;
  --surface: #ffffff;
  --surface-subtle: #f8f9fb;
  --surface-muted: #f1f3f6;
  --surface-hover: #edf3f7;
  --surface-active: #e7f1f8;
  --border: #e3e7ec;
  --border-strong: #cfd6de;
  --text: #17202c;
  --text-secondary: #465466;
  --text-tertiary: #758195;
  --accent: #146b9f;
  --accent-strong: #0e5784;
  --accent-soft: #e7f2f8;
  --success: #25805a;
  --warning: #9a5a16;
  --danger: #b3424c;
  --topbar-height: 64px;
  --rail-width: 76px;
  --stream-width: 400px;
  font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
  font-synthesis: none;
  text-rendering: optimizeLegibility;
}

* {
  box-sizing: border-box;
}

html,
body {
  width: 100%;
  min-width: 320px;
  height: 100%;
  margin: 0;
}

body {
  background: var(--canvas);
  color: var(--text);
}

button,
input {
  font: inherit;
}

button,
a {
  -webkit-tap-highlight-color: transparent;
}

button {
  color: inherit;
}

button:focus-visible,
input:focus-visible,
a:focus-visible {
  outline: 2px solid var(--accent);
  outline-offset: 2px;
}

[hidden] {
  display: none !important;
}

.sr-only {
  position: absolute !important;
  width: 1px !important;
  height: 1px !important;
  padding: 0 !important;
  overflow: hidden !important;
  clip: rect(0, 0, 0, 0) !important;
  white-space: nowrap !important;
  border: 0 !important;
}

.icon {
  display: block;
  width: 18px;
  height: 18px;
  flex: 0 0 auto;
}

.eyebrow {
  margin: 0;
  color: var(--text-tertiary);
  font-size: 10px;
  font-weight: 750;
  letter-spacing: 0;
  line-height: 1.2;
}

.app-body {
  overflow: hidden;
}

.app-shell {
  display: grid;
  grid-template-rows: var(--topbar-height) minmax(0, 1fr);
  width: 100%;
  height: 100dvh;
  overflow: hidden;
  background: var(--surface);
}

.topbar {
  position: relative;
  z-index: 20;
  display: grid;
  grid-template-columns: 124px minmax(240px, 300px) minmax(260px, 680px) 84px;
  align-items: center;
  gap: 16px;
  height: var(--topbar-height);
  padding: 0 18px;
  border-bottom: 1px solid var(--border);
  background: rgba(255, 255, 255, 0.96);
  box-shadow: 0 1px 0 rgba(20, 29, 40, 0.02);
}

.brand,
.login-brand {
  display: flex;
  align-items: center;
  gap: 10px;
  min-width: 0;
}

.brand-copy,
.login-brand > div {
  display: grid;
  min-width: 0;
}

.brand-copy strong {
  font-size: 15px;
  font-weight: 760;
  line-height: 1.1;
}

.brand-copy span {
  margin-top: 2px;
  color: var(--text-tertiary);
  font-size: 10px;
}

.brand-mark {
  display: inline-grid;
  flex: 0 0 auto;
  width: 32px;
  height: 32px;
  place-items: center;
  border-radius: 7px;
  background: var(--accent);
  color: #ffffff;
  box-shadow: inset -7px 0 0 rgba(255, 255, 255, 0.1);
}

.brand-mark .icon {
  width: 17px;
  height: 17px;
  stroke-width: 2;
}

.account-switcher {
  position: relative;
  min-width: 0;
}

.account-trigger {
  display: grid;
  grid-template-columns: 32px minmax(0, 1fr) 18px;
  align-items: center;
  gap: 10px;
  width: 100%;
  min-height: 44px;
  padding: 5px 9px;
  border: 1px solid transparent;
  border-radius: 7px;
  background: transparent;
  cursor: pointer;
  text-align: left;
  transition: background-color 140ms ease, border-color 140ms ease;
}

.account-trigger:hover,
.account-trigger[aria-expanded="true"] {
  border-color: var(--border);
  background: var(--surface-subtle);
}

.account-avatar,
.message-avatar {
  display: grid;
  place-items: center;
  border-radius: 50%;
  background: #dceaf3;
  color: #15567f;
  font-size: 12px;
  font-weight: 760;
}

.account-avatar {
  width: 32px;
  height: 32px;
}

.account-copy {
  display: grid;
  min-width: 0;
}

.account-copy strong,
.account-copy span {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.account-copy strong {
  font-size: 13px;
  font-weight: 700;
}

.account-copy span {
  margin-top: 1px;
  color: var(--text-tertiary);
  font-size: 10px;
}

.chevron {
  color: var(--text-tertiary);
  transition: transform 140ms ease;
}

.chevron .icon {
  width: 16px;
  height: 16px;
}

.account-trigger[aria-expanded="true"] .chevron {
  transform: rotate(180deg);
}

.account-menu {
  position: absolute;
  top: calc(100% + 7px);
  left: 0;
  z-index: 30;
  width: min(380px, calc(100vw - 24px));
  max-height: min(440px, calc(100dvh - 82px));
  overflow: auto;
  padding: 6px;
  border: 1px solid var(--border-strong);
  border-radius: 8px;
  background: var(--surface);
  box-shadow: 0 18px 44px rgba(30, 42, 57, 0.16);
  animation: menu-enter 140ms ease-out both;
}

@keyframes menu-enter {
  from { opacity: 0; transform: translateY(-4px); }
  to { opacity: 1; transform: translateY(0); }
}

.account-option {
  display: grid;
  grid-template-columns: 32px minmax(0, 1fr) auto;
  align-items: center;
  gap: 10px;
  width: 100%;
  min-height: 62px;
  padding: 8px 9px;
  border: 0;
  border-radius: 6px;
  background: transparent;
  cursor: pointer;
  text-align: left;
  transition: background-color 120ms ease;
}

.account-option:hover,
.account-option.is-active {
  background: var(--surface-hover);
}

.account-option-copy {
  display: grid;
  min-width: 0;
}

.account-option-copy strong,
.account-option-copy span {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.account-option-copy strong {
  font-size: 13px;
}

.account-option-copy span,
.account-route {
  margin-top: 2px;
  color: var(--text-tertiary);
  font-size: 10px;
}

.account-empty {
  display: grid;
  gap: 4px;
  padding: 14px 12px;
}

.account-empty strong {
  font-size: 12px;
}

.account-empty span {
  color: var(--text-tertiary);
  font-size: 10px;
  line-height: 1.5;
}

.account-state {
  width: 7px;
  height: 7px;
  border-radius: 50%;
  background: var(--success);
}

.account-state.is-error {
  background: var(--danger);
}

.global-search {
  display: grid;
  grid-template-columns: 18px minmax(0, 1fr);
  align-items: center;
  gap: 9px;
  width: 100%;
  min-height: 38px;
  padding: 0 12px;
  border: 1px solid var(--border);
  border-radius: 7px;
  background: var(--surface-muted);
  color: var(--text-tertiary);
  transition: border-color 140ms ease, background-color 140ms ease, box-shadow 140ms ease;
}

.global-search:focus-within {
  border-color: #8dbbd5;
  background: var(--surface);
  box-shadow: 0 0 0 3px rgba(20, 107, 159, 0.08);
}

.global-search .icon {
  width: 16px;
  height: 16px;
}

.global-search input {
  width: 100%;
  min-width: 0;
  border: 0;
  outline: 0;
  background: transparent;
  color: var(--text);
  font-size: 12px;
}

.global-search input::placeholder {
  color: #8a94a2;
}

.topbar-actions {
  display: flex;
  align-items: center;
  justify-content: flex-end;
  gap: 4px;
}

.topbar-actions form {
  margin: 0;
}

.icon-button {
  display: inline-grid;
  width: 36px;
  height: 36px;
  place-items: center;
  padding: 0;
  border: 1px solid transparent;
  border-radius: 7px;
  background: transparent;
  color: var(--text-secondary);
  cursor: pointer;
  transition: background-color 140ms ease, border-color 140ms ease, color 140ms ease;
}

.icon-button:hover {
  border-color: var(--border);
  background: var(--surface-muted);
  color: var(--text);
}

.icon-button:disabled {
  cursor: default;
  opacity: 0.55;
}

.icon-button.is-spinning .icon {
  animation: rotate 700ms linear infinite;
}

@keyframes rotate {
  to { transform: rotate(360deg); }
}

.workspace {
  display: grid;
  grid-template-columns: var(--rail-width) minmax(340px, var(--stream-width)) minmax(0, 1fr);
  min-height: 0;
  overflow: hidden;
}

.folder-rail {
  display: flex;
  flex-direction: column;
  min-height: 0;
  padding: 12px 8px 14px;
  border-right: 1px solid var(--border);
  background: var(--surface-subtle);
}

.folder-nav {
  display: grid;
  gap: 6px;
}

.folder-button {
  position: relative;
  display: grid;
  width: 60px;
  min-height: 58px;
  place-items: center;
  gap: 4px;
  padding: 7px 4px;
  border: 0;
  border-radius: 7px;
  background: transparent;
  color: var(--text-tertiary);
  cursor: pointer;
  transition: background-color 140ms ease, color 140ms ease;
}

.folder-button::before {
  content: "";
  position: absolute;
  top: 12px;
  bottom: 12px;
  left: -8px;
  width: 3px;
  border-radius: 0 3px 3px 0;
  background: transparent;
}

.folder-button .icon {
  width: 19px;
  height: 19px;
}

.folder-button small {
  font-size: 10px;
  line-height: 1.1;
  white-space: nowrap;
}

.folder-button:hover {
  background: var(--surface-muted);
  color: var(--text);
}

.folder-button.is-active {
  background: var(--accent-soft);
  color: var(--accent);
}

.folder-button.is-active::before {
  background: var(--accent);
}

.rail-status {
  display: grid;
  justify-items: center;
  gap: 6px;
  margin-top: auto;
  color: var(--text-tertiary);
  font-size: 9px;
}

.connection-state,
.status-dot {
  width: 8px;
  height: 8px;
  border-radius: 50%;
  background: var(--text-tertiary);
}

.connection-state.is-online,
.status-dot {
  background: var(--success);
  box-shadow: 0 0 0 3px rgba(37, 128, 90, 0.11);
}

.connection-state.is-error {
  background: var(--danger);
  box-shadow: 0 0 0 3px rgba(179, 66, 76, 0.11);
}

.stream-pane,
.reader-pane {
  min-width: 0;
  min-height: 0;
  overflow: hidden;
}

.stream-pane {
  display: grid;
  grid-template-rows: auto auto auto minmax(0, 1fr) auto;
  border-right: 1px solid var(--border);
  background: var(--surface);
}

.stream-header {
  display: flex;
  align-items: end;
  justify-content: space-between;
  gap: 18px;
  min-height: 84px;
  padding: 18px 18px 13px;
}

.stream-header h1 {
  margin: 4px 0 0;
  font-size: 20px;
  font-weight: 760;
  line-height: 1.15;
}

.stream-summary {
  display: grid;
  justify-items: end;
  gap: 3px;
  min-width: 0;
  color: var(--text-tertiary);
  font-size: 9px;
  text-align: right;
}

.stream-summary span {
  max-width: 150px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.message-count {
  color: var(--text-secondary);
  font-size: 11px;
  font-weight: 650;
}

.filter-bar {
  display: flex;
  gap: 3px;
  min-height: 44px;
  padding: 4px 14px 10px;
  border-bottom: 1px solid var(--border);
}

.filter-button {
  height: 30px;
  padding: 0 11px;
  border: 1px solid transparent;
  border-radius: 6px;
  background: transparent;
  color: var(--text-tertiary);
  cursor: pointer;
  font-size: 11px;
  transition: background-color 120ms ease, border-color 120ms ease, color 120ms ease;
}

.filter-button:hover {
  background: var(--surface-muted);
  color: var(--text);
}

.filter-button.is-active {
  border-color: #cce0ec;
  background: var(--accent-soft);
  color: var(--accent);
  font-weight: 700;
}

.stream-notice {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  padding: 10px 14px;
  border-bottom: 1px solid #efcdd0;
  background: #fff7f7;
  color: #8e3139;
  font-size: 11px;
  line-height: 1.45;
}

.notice-copy {
  min-width: 0;
  overflow-wrap: anywhere;
}

.stream-notice button {
  flex: 0 0 auto;
  min-height: 28px;
  padding: 0 9px;
  border: 1px solid #dfb5b9;
  border-radius: 5px;
  background: var(--surface);
  color: #8e3139;
  cursor: pointer;
  font-size: 10px;
}

.message-stage {
  position: relative;
  min-height: 0;
  overflow: hidden;
}

.message-list {
  height: 100%;
  min-height: 0;
  overflow-x: hidden;
  overflow-y: auto;
  overscroll-behavior: contain;
  scrollbar-gutter: stable;
}

.message-item {
  position: relative;
  display: grid;
  grid-template-columns: 34px minmax(0, 1fr) auto;
  grid-template-rows: 19px 22px 18px 16px;
  gap: 0 10px;
  width: 100%;
  height: 94px;
  padding: 9px 14px 8px;
  border: 0;
  border-bottom: 1px solid var(--border);
  background: var(--surface);
  cursor: pointer;
  text-align: left;
  content-visibility: auto;
  contain-intrinsic-size: 94px;
  contain: layout paint style;
  transition: background-color 120ms ease;
}

.message-item::before {
  content: "";
  position: absolute;
  top: 0;
  bottom: 0;
  left: 0;
  width: 3px;
  background: transparent;
}

.message-item:hover {
  background: var(--surface-hover);
}

.message-item.is-active {
  background: var(--surface-active);
}

.message-item.is-active::before {
  background: var(--accent);
}

.message-item[hidden] {
  display: none;
}

.message-avatar {
  grid-row: 1 / span 2;
  width: 32px;
  height: 32px;
  margin-top: 2px;
}

.message-sender,
.message-subject,
.message-preview,
.message-code {
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.message-sender {
  align-self: center;
  font-size: 12px;
  font-weight: 720;
}

.message-time {
  align-self: center;
  color: var(--text-tertiary);
  font-size: 10px;
  white-space: nowrap;
}

.message-subject {
  grid-column: 2 / 4;
  align-self: center;
  color: var(--text-secondary);
  font-size: 12px;
  font-weight: 590;
}

.message-preview {
  grid-column: 2 / 4;
  align-self: center;
  color: var(--text-tertiary);
  font-size: 10px;
}

.message-meta {
  grid-column: 2 / 4;
  display: flex;
  align-items: center;
  gap: 7px;
  min-width: 0;
  color: var(--text-tertiary);
  font-size: 9px;
}

.message-code {
  color: var(--accent);
  font-weight: 760;
  font-variant-numeric: tabular-nums;
}

.attachment-dot {
  color: var(--warning);
}

.stream-empty {
  position: absolute;
  inset: 0;
  display: grid;
  align-content: center;
  justify-items: center;
  gap: 6px;
  padding: 30px 22px;
  background: var(--surface);
  color: var(--text-tertiary);
  text-align: center;
}

.empty-mark,
.placeholder-mark {
  display: grid;
  place-items: center;
  border: 1px solid var(--border);
  border-radius: 8px;
  background: var(--surface-subtle);
  color: var(--accent);
}

.empty-mark {
  width: 42px;
  height: 42px;
  margin-bottom: 7px;
}

.empty-mark .icon {
  width: 19px;
  height: 19px;
}

.stream-empty strong {
  color: var(--text-secondary);
  font-size: 13px;
}

.stream-empty > span:not(.empty-mark) {
  max-width: 260px;
  font-size: 11px;
  line-height: 1.5;
}

.stream-empty button {
  min-height: 30px;
  margin-top: 6px;
  padding: 0 11px;
  border: 1px solid var(--border-strong);
  border-radius: 6px;
  background: var(--surface);
  color: var(--text-secondary);
  cursor: pointer;
  font-size: 10px;
}

.stream-empty button:hover {
  background: var(--surface-muted);
}

.pagination {
  display: grid;
  grid-template-columns: 34px 1fr 34px;
  align-items: center;
  gap: 8px;
  min-height: 48px;
  padding: 7px 14px;
  border-top: 1px solid var(--border);
  background: var(--surface);
}

.pagination button {
  display: grid;
  width: 32px;
  height: 32px;
  place-items: center;
  padding: 0;
  border: 1px solid var(--border);
  border-radius: 6px;
  background: var(--surface);
  color: var(--text-secondary);
  cursor: pointer;
  font-size: 14px;
}

.pagination button:last-child {
  justify-self: end;
}

.pagination button:hover:not(:disabled) {
  border-color: var(--border-strong);
  background: var(--surface-muted);
}

.pagination button:disabled {
  cursor: default;
  opacity: 0.35;
}

.pagination span {
  color: var(--text-tertiary);
  font-size: 10px;
  text-align: center;
  white-space: nowrap;
}

.reader-pane {
  position: relative;
  overflow-y: auto;
  overscroll-behavior: contain;
  scrollbar-gutter: stable;
  background: var(--surface-subtle);
}

.reader-mobile-bar {
  display: none;
}

.reader-content {
  min-height: 100%;
}

.reader-placeholder {
  display: grid;
  min-height: 100%;
  place-content: center;
  justify-items: center;
  gap: 7px;
  padding: 34px;
  color: var(--text-tertiary);
  text-align: center;
}

.placeholder-mark {
  width: 48px;
  height: 48px;
  margin-bottom: 7px;
}

.placeholder-mark .icon {
  width: 21px;
  height: 21px;
}

.reader-placeholder h2 {
  margin: 0;
  color: var(--text-secondary);
  font-size: 16px;
}

.reader-placeholder p {
  max-width: 360px;
  margin: 0;
  font-size: 11px;
  line-height: 1.6;
}

.reader-document {
  width: 100%;
  min-height: 100%;
  padding-bottom: 48px;
  background: var(--surface);
  animation: reader-enter 160ms ease-out both;
}

@keyframes reader-enter {
  from { opacity: 0; transform: translateY(4px); }
  to { opacity: 1; transform: translateY(0); }
}

.reader-header,
.code-banner,
.attachment-section,
.reader-body {
  width: calc(100% - 56px);
  max-width: 920px;
  margin-right: auto;
  margin-left: auto;
}

.reader-header {
  padding: 36px 0 24px;
  border-bottom: 1px solid var(--border);
}

.reader-header-row {
  display: flex;
  align-items: start;
  justify-content: space-between;
  gap: 24px;
}

.reader-header h2 {
  margin: 5px 0 12px;
  color: var(--text);
  font-size: 27px;
  font-weight: 750;
  line-height: 1.25;
  overflow-wrap: anywhere;
}

.reader-open {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  flex: 0 0 auto;
  min-height: 34px;
  padding: 8px 11px;
  border: 1px solid var(--border-strong);
  border-radius: 6px;
  background: var(--surface);
  color: var(--accent);
  font-size: 10px;
  font-weight: 680;
  text-decoration: none;
  transition: background-color 120ms ease, border-color 120ms ease;
}

.reader-open .icon {
  width: 14px;
  height: 14px;
}

.reader-open:hover {
  border-color: #b9d4e4;
  background: var(--accent-soft);
}

.reader-sender {
  display: flex;
  flex-wrap: wrap;
  gap: 5px 12px;
  color: var(--text-secondary);
  font-size: 11px;
  line-height: 1.5;
}

.reader-sender span {
  color: var(--text-tertiary);
}

.code-banner {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 18px;
  margin-top: 18px;
  padding: 14px 16px;
  border: 1px solid #bdd8e7;
  border-left: 3px solid var(--accent);
  border-radius: 7px;
  background: #eef7fb;
}

.code-copy {
  display: grid;
  gap: 3px;
}

.code-copy span {
  color: var(--text-secondary);
  font-size: 10px;
}

.code-copy strong {
  color: var(--accent-strong);
  font-size: 25px;
  font-weight: 780;
  font-variant-numeric: tabular-nums;
  letter-spacing: 0;
}

.copy-code {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  min-height: 32px;
  padding: 0 11px;
  border: 1px solid #9fc8dc;
  border-radius: 5px;
  background: var(--surface);
  color: var(--accent);
  cursor: pointer;
  font-size: 10px;
}

.copy-code .icon {
  width: 14px;
  height: 14px;
}

.copy-code:hover {
  background: var(--accent-soft);
}

.reader-body {
  padding: 26px 0 24px;
}

.mail-body-frame {
  display: block;
  width: 100%;
  min-height: max(620px, calc(100dvh - 290px));
  border: 0;
  background: var(--surface);
}

.mail-body-text {
  max-width: 76ch;
  margin: 0;
  color: var(--text);
  font: 14px/1.78 Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}

.attachment-section {
  margin-top: 18px;
  padding: 16px;
  border-top: 1px solid var(--border);
  border-bottom: 1px solid var(--border);
  background: var(--surface-subtle);
}

.attachment-section h3 {
  margin: 0 0 9px;
  font-size: 10px;
}

.attachment-list {
  display: flex;
  flex-wrap: wrap;
  gap: 7px;
}

.attachment-item {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  min-height: 32px;
  padding: 5px 9px;
  border: 1px solid var(--border);
  border-radius: 5px;
  background: var(--surface);
  color: var(--text-secondary);
  font-size: 10px;
}

.attachment-item .icon {
  width: 14px;
  height: 14px;
}

.skeleton-list {
  padding: 0 14px;
}

.skeleton-row {
  display: grid;
  grid-template-columns: 32px 1fr;
  gap: 10px;
  height: 94px;
  align-items: center;
  border-bottom: 1px solid var(--border);
}

.skeleton-avatar,
.skeleton-line {
  background: #e7eaee;
  animation: pulse 1.05s ease-in-out infinite alternate;
}

.skeleton-avatar {
  width: 32px;
  height: 32px;
  border-radius: 50%;
}

.skeleton-lines {
  display: grid;
  gap: 8px;
}

.skeleton-line {
  height: 8px;
  border-radius: 4px;
}

.skeleton-line:last-child {
  width: 68%;
}

.reader-skeleton {
  display: grid;
  width: calc(100% - 56px);
  max-width: 920px;
  gap: 14px;
  margin: 0 auto;
  padding: 40px 0;
}

.reader-skeleton .skeleton-line:nth-child(1) { width: 24%; }
.reader-skeleton .skeleton-line:nth-child(2) { width: 74%; height: 22px; }
.reader-skeleton .skeleton-line:nth-child(3) { width: 46%; }
.reader-skeleton .skeleton-line:nth-child(n + 4) { width: 100%; }

@keyframes pulse {
  from { opacity: 0.5; }
  to { opacity: 1; }
}

.toast {
  position: fixed;
  right: 18px;
  bottom: 18px;
  z-index: 60;
  max-width: min(360px, calc(100vw - 36px));
  padding: 10px 13px;
  border: 1px solid #313c49;
  border-radius: 7px;
  background: #202a35;
  color: #ffffff;
  box-shadow: 0 12px 34px rgba(23, 32, 44, 0.2);
  font-size: 11px;
  animation: toast-enter 150ms ease-out both;
}

@keyframes toast-enter {
  from { opacity: 0; transform: translateY(6px); }
  to { opacity: 1; transform: translateY(0); }
}

.login-body {
  overflow: auto;
}

.login-page {
  display: grid;
  grid-template-columns: minmax(340px, 44%) minmax(0, 1fr);
  width: 100%;
  min-height: 100dvh;
  background: var(--surface);
}

.login-intro {
  display: flex;
  flex-direction: column;
  min-height: 100dvh;
  padding: 42px 48px;
  background: #202326;
  color: #ffffff;
}

.login-brand .brand-mark-large {
  width: 38px;
  height: 38px;
  background: #1b7caf;
}

.login-brand .eyebrow {
  color: #a9b0b4;
}

.login-brand strong {
  margin-top: 2px;
  font-size: 16px;
}

.login-intro-copy {
  max-width: 500px;
  margin: auto 0;
  padding: 48px 0;
}

.login-kicker {
  margin: 0 0 14px;
  color: #abb3b7;
  font-size: 12px;
  font-weight: 650;
}

.login-intro-copy h1 {
  margin: 0;
  color: #ffffff;
  font-size: 52px;
  font-weight: 720;
  line-height: 1.12;
  letter-spacing: 0;
}

.login-intro-copy > p:last-child {
  max-width: 420px;
  margin: 20px 0 0;
  color: #c2c7ca;
  font-size: 14px;
  line-height: 1.7;
}

.login-intro-foot {
  margin: 0;
  color: #90999e;
  font-size: 10px;
}

.login-intro > * {
  animation: login-enter 400ms ease-out both;
}

.login-intro-copy {
  animation-delay: 70ms;
}

.login-intro-foot {
  animation-delay: 130ms;
}

@keyframes login-enter {
  from { opacity: 0; transform: translateY(8px); }
  to { opacity: 1; transform: translateY(0); }
}

.login-form-pane {
  display: grid;
  min-height: 100dvh;
  place-items: center;
  padding: 42px;
  background: var(--surface);
}

.login-form-wrap {
  width: min(100%, 390px);
}

.login-form-heading {
  margin-bottom: 30px;
}

.login-form-heading h2 {
  margin: 8px 0 8px;
  font-size: 28px;
  font-weight: 740;
  line-height: 1.2;
}

.login-form-heading p,
.login-security-note {
  margin: 0;
  color: var(--text-tertiary);
  font-size: 11px;
  line-height: 1.55;
}

.login-form {
  display: grid;
  gap: 9px;
}

.login-form label {
  color: var(--text-secondary);
  font-size: 11px;
  font-weight: 680;
}

.login-form input {
  width: 100%;
  height: 46px;
  padding: 0 13px;
  border: 1px solid var(--border-strong);
  border-radius: 6px;
  background: var(--surface);
  color: var(--text);
  outline: 0;
  transition: border-color 140ms ease, box-shadow 140ms ease;
}

.login-form input:focus {
  border-color: var(--accent);
  box-shadow: 0 0 0 3px rgba(20, 107, 159, 0.1);
}

.login-form button {
  display: flex;
  align-items: center;
  justify-content: space-between;
  height: 46px;
  margin-top: 7px;
  padding: 0 15px;
  border: 0;
  border-radius: 6px;
  background: var(--accent);
  color: #ffffff;
  cursor: pointer;
  font-size: 12px;
  font-weight: 700;
  transition: background-color 140ms ease, transform 140ms ease;
}

.login-form button:hover {
  background: var(--accent-strong);
}

.login-form button:active {
  transform: translateY(1px);
}

.login-alert {
  margin: 0 0 18px;
  padding: 11px 12px;
  border: 1px solid #ecc8cb;
  border-radius: 6px;
  background: #fff6f6;
  color: #8c3038;
  font-size: 11px;
  line-height: 1.55;
}

.login-security-note {
  display: flex;
  align-items: center;
  gap: 9px;
  margin-top: 22px;
}

.login-security-note .status-dot {
  width: 7px;
  height: 7px;
  box-shadow: none;
}

@media (max-width: 1180px) {
  :root {
    --stream-width: 360px;
  }

  .topbar {
    grid-template-columns: 46px minmax(220px, 280px) minmax(220px, 1fr) 84px;
    gap: 12px;
  }

  .brand-copy {
    display: none;
  }

  .reader-header,
  .code-banner,
  .attachment-section,
  .reader-body,
  .reader-skeleton {
    width: calc(100% - 44px);
  }
}

@media (max-width: 900px) {
  :root {
    --rail-width: 68px;
    --stream-width: 336px;
  }

  .topbar {
    padding: 0 12px;
  }

  .folder-rail {
    padding-right: 4px;
    padding-left: 4px;
  }

  .folder-button {
    width: 60px;
  }

  .reader-header-row {
    display: grid;
  }

  .reader-open {
    width: fit-content;
  }

  .reader-header h2 {
    font-size: 23px;
  }
}

@media (max-width: 760px) {
  :root {
    --topbar-height: auto;
  }

  .app-shell {
    grid-template-rows: auto minmax(0, 1fr);
  }

  .topbar {
    grid-template-areas:
      "brand actions"
      "account account"
      "search search";
    grid-template-columns: minmax(0, 1fr) auto;
    gap: 8px 10px;
    height: auto;
    padding: calc(9px + env(safe-area-inset-top)) 12px 10px;
  }

  .brand {
    grid-area: brand;
  }

  .brand-copy {
    display: grid;
  }

  .account-switcher {
    grid-area: account;
  }

  .account-trigger {
    min-height: 42px;
    padding-right: 7px;
    padding-left: 7px;
    border-color: var(--border);
    background: var(--surface-subtle);
  }

  .account-menu {
    width: 100%;
    max-height: min(360px, calc(100dvh - 168px));
  }

  .global-search {
    grid-area: search;
    min-height: 40px;
  }

  .topbar-actions {
    grid-area: actions;
  }

  .workspace {
    position: relative;
    grid-template-columns: 1fr;
    grid-template-rows: 54px minmax(0, 1fr);
  }

  .folder-rail {
    grid-row: 1;
    display: flex;
    flex-direction: row;
    align-items: center;
    padding: 5px 10px;
    border-right: 0;
    border-bottom: 1px solid var(--border);
  }

  .folder-nav {
    display: flex;
    gap: 4px;
  }

  .folder-button {
    display: flex;
    width: auto;
    min-height: 40px;
    gap: 7px;
    padding: 0 11px;
  }

  .folder-button::before {
    top: auto;
    right: 9px;
    bottom: -5px;
    left: 9px;
    width: auto;
    height: 2px;
    border-radius: 2px 2px 0 0;
  }

  .folder-button .icon {
    width: 17px;
    height: 17px;
  }

  .rail-status {
    display: flex;
    align-items: center;
    gap: 7px;
    margin: 0 2px 0 auto;
  }

  .stream-pane {
    grid-row: 2;
    border-right: 0;
  }

  .stream-header {
    min-height: 76px;
    padding: 15px 16px 11px;
  }

  .filter-bar {
    min-height: 42px;
    padding-right: 12px;
    padding-left: 12px;
  }

  .message-item {
    height: 96px;
    contain-intrinsic-size: 96px;
  }

  .reader-pane {
    position: absolute;
    inset: 0;
    z-index: 12;
    display: none;
    overflow-y: auto;
    background: var(--surface);
  }

  .app-shell[data-mobile-view="reader"] .reader-pane {
    display: block;
    animation: mobile-reader-enter 180ms ease-out both;
  }

  @keyframes mobile-reader-enter {
    from { opacity: 0; transform: translateX(10px); }
    to { opacity: 1; transform: translateX(0); }
  }

  .reader-mobile-bar {
    position: sticky;
    top: 0;
    z-index: 8;
    display: grid;
    grid-template-columns: auto minmax(0, 1fr) auto;
    align-items: center;
    min-height: 52px;
    padding: 0 12px;
    border-bottom: 1px solid var(--border);
    background: rgba(255, 255, 255, 0.96);
  }

  .reader-back {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    min-height: 36px;
    padding: 0 7px 0 2px;
    border: 0;
    border-radius: 6px;
    background: transparent;
    color: var(--accent);
    cursor: pointer;
    font-size: 11px;
  }

  .reader-back .icon {
    width: 17px;
    height: 17px;
  }

  .reader-mobile-bar strong {
    grid-column: 2;
    overflow: hidden;
    color: var(--text-secondary);
    font-size: 11px;
    text-align: center;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .reader-header,
  .code-banner,
  .attachment-section,
  .reader-body,
  .reader-skeleton {
    width: calc(100% - 32px);
  }

  .reader-header {
    padding: 24px 0 19px;
  }

  .reader-header h2 {
    font-size: 21px;
  }

  .code-banner {
    margin-top: 14px;
    padding: 12px 13px;
  }

  .code-copy strong {
    font-size: 23px;
  }

  .reader-body {
    padding-top: 22px;
  }

  .mail-body-frame {
    min-height: calc(100dvh - 250px);
  }

  .reader-document {
    padding-bottom: max(34px, env(safe-area-inset-bottom));
  }

  .toast {
    right: 12px;
    bottom: calc(12px + env(safe-area-inset-bottom));
    max-width: calc(100vw - 24px);
  }
}

@media (max-width: 700px) {
  .login-page {
    grid-template-columns: 1fr;
    grid-template-rows: auto minmax(0, 1fr);
  }

  .login-intro {
    min-height: 220px;
    padding: calc(26px + env(safe-area-inset-top)) 24px 24px;
  }

  .login-intro-copy {
    margin: 30px 0 0;
    padding: 0;
  }

  .login-intro-copy h1 {
    font-size: 34px;
  }

  .login-intro-copy > p:last-child {
    margin-top: 12px;
    font-size: 12px;
  }

  .login-intro-foot {
    display: none;
  }

  .login-form-pane {
    min-height: 0;
    place-items: start center;
    padding: 34px 24px max(34px, env(safe-area-inset-bottom));
  }

  .login-form-wrap {
    width: 100%;
  }
}

@media (max-width: 420px) {
  .brand-copy span {
    display: none;
  }

  .folder-button {
    padding-right: 9px;
    padding-left: 9px;
  }

  .folder-button small {
    font-size: 9px;
  }

  .stream-summary span {
    max-width: 120px;
  }
}

@media (prefers-reduced-motion: reduce) {
  *,
  *::before,
  *::after {
    scroll-behavior: auto !important;
    animation-duration: 0.01ms !important;
    animation-iteration-count: 1 !important;
    transition-duration: 0.01ms !important;
  }
}
`;
