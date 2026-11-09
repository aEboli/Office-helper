using System;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using ChatSheet.AddIn;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json.Linq;

namespace ChatSheet.PaneHarness
{
    internal static class WorkBuddyUiTests
    {
        internal static async Task<CoreWebView2> WaitForCoreAsync(TaskPaneControl pane)
        {
            var field = typeof(TaskPaneControl).GetField(
                "_webView",
                BindingFlags.NonPublic | BindingFlags.Instance);
            if (field == null)
            {
                throw new MissingFieldException(typeof(TaskPaneControl).FullName, "_webView");
            }

            var web = field.GetValue(pane) as WebView2;
            if (web == null)
            {
                throw new InvalidOperationException("PaneHarness 无法取得 WebView2 控件");
            }

            for (var i = 0; i < 120; i++)
            {
                if (web.CoreWebView2 != null)
                {
                    return web.CoreWebView2;
                }
                await Task.Delay(250);
            }

            var fallbackField = typeof(TaskPaneControl).GetField(
                "_fallback",
                BindingFlags.NonPublic | BindingFlags.Instance);
            var fallback = fallbackField?.GetValue(pane) as Label;
            throw new TimeoutException(
                "WebView2 30 秒内未初始化" +
                (string.IsNullOrWhiteSpace(fallback?.Text) ? string.Empty : "：" + fallback.Text));
        }

        internal static async Task PreparePickerFixtureAsync(TaskPaneControl pane)
        {
            var core = await WaitForCoreAsync(pane);
            await core.ExecuteScriptAsync(@"import('./scripts/bridge.js').then(({request}) => request('settings.get')).then(s => {
              const channel=window.chrome.webview, original=channel.postMessage.bind(channel);
              const models=['seed-ok','seed-bad','seed-unknown','deepseek/deepseek-v4-flash-vision-preview'];
              const settings={...s, mode:'CustomApi', customProtocol:'openai-chat-completions',
                customBaseUrl:'https://seed.example.test/v1', model:'seed-ok', effectiveModel:'seed-ok',
                thinking:'High', thinkingSupported:['Off','Minimal','Low','Medium','High'],
                favorites:[], availability:{'seed-ok':'Available','seed-bad':'Unavailable'}, onlyFavoriteModels:false};
              channel.postMessage=m => {
                if(m.channel==='settings.get' || m.channel==='models.list') {
                  queueMicrotask(()=>channel.dispatchEvent(new MessageEvent('message', {data:{
                    kind:'response',id:m.id,ok:true,data:m.channel==='settings.get'?settings:{models}
                  }})));
                } else { original(m); }
              };
              window.__pickerFixtureReady=true;
            });");
            for (var i = 0; i < 120; i++)
            {
                if (await core.ExecuteScriptAsync("!!window.__pickerFixtureReady") == "true") { return; }
                await Task.Delay(250);
            }
            throw new TimeoutException("Picker fixture settings did not load");
        }

        internal static int Run(int width, string captureDirectory)
        {
            var failed = 0;
            void Check(bool ok, string label)
            {
                Console.WriteLine((ok ? "PASS " : "FAIL ") + label);
                if (!ok) { failed++; }
            }

            using (var form = new Form { ClientSize = new System.Drawing.Size(width, 760), ShowInTaskbar = false })
            using (var pane = new TaskPaneControl { Dock = DockStyle.Fill })
            {
                form.Controls.Add(pane);
                form.Shown += async (sender, args) =>
                {
                    try
                    {
                        for (var i = 0; i < 60 && !pane.ReadThemeState().StartsWith("theme="); i++)
                        {
                            await Task.Delay(250);
                        }
                        var core = await WaitForCoreAsync(pane);
                        Check(await core.ExecuteScriptAsync($"innerWidth === {width}") == "true", "requested panel viewport width");
                        // 等真实授权目录与初始渲染完成，不向用户设置写入测试目录。
                        await core.ExecuteScriptAsync("import('./scripts/bridge.js').then(({request}) => request('settings.get')).then(s => {window.__wbSettings=s;});");
                        for (var i = 0; i < 120; i++)
                        {
                            if (await core.ExecuteScriptAsync("!!window.__wbSettings") == "true") { break; }
                            await Task.Delay(250);
                        }
                        var authorizationStatus = await core.ExecuteScriptAsync("window.__wbSettings?.authorization?.status || ''");
                        var connectionMode = await core.ExecuteScriptAsync("window.__wbSettings?.mode || ''");
                        Check(authorizationStatus == "\"authorized\"" || authorizationStatus == "\"unauthorized\"" || authorizationStatus == "\"unavailable\"",
                            "authorization state is explicit");
                        // 版本号每次发版都会变，写死常量必然滞后（原先是 0.10.3.13，装的是 0.10.3.19）；
                        // 改为读页面自己的 data-version，只断言「无前缀四段构建号」且文本与 data-version 一致。
                        Check(await core.ExecuteScriptAsync("(() => { const n = document.getElementById('app-version'); const v = n && n.dataset.version; return !!v && /^\\d+\\.\\d+\\.\\d+\\.\\d+$/.test(v) && n.textContent === v; })()") == "true", "version label is a bare four-part build number equal to data-version");
                        if (authorizationStatus == "\"authorized\"")
                        {
                            await core.ExecuteScriptAsync("document.getElementById('picker-trigger').click()");
                            for (var i = 0; i < 40; i++)
                            {
                                if (await core.ExecuteScriptAsync("document.querySelectorAll('.picker-item-multiplier').length > 0") == "true") { break; }
                                await Task.Delay(250);
                            }
                            var result = JObject.Parse(await core.ExecuteScriptAsync(@"(() => {
                          const nodes = [...document.querySelectorAll('.picker-item-multiplier')];
                          const bounds = document.getElementById('picker-pop').getBoundingClientRect();
                          return {
                            count: nodes.length,
                            modelCount: document.querySelectorAll('.picker-item').length,
                            metadataMatches: [...document.querySelectorAll('.picker-item')].every(row => {
                              const id = row.querySelector('.picker-item-name')?.textContent;
                              const model = window.__wbSettings.authorization.models.find(m => m.modelId === id);
                              return (row.querySelector('.picker-item-multiplier')?.textContent || '') === (model?.multiplier || '');
                            }),
                            visible: nodes.every(n => { const a=n.getBoundingClientRect(), b=n.previousElementSibling.getBoundingClientRect(); return a.width>0 && a.right<=innerWidth && a.left>=b.right-1; }),
                            fits: bounds.left>=0 && bounds.right<=innerWidth && bounds.top>=0,
                            overflow: document.documentElement.scrollWidth>innerWidth,
                          };
                        })()"));
                            Check(result.Value<int>("modelCount") > 0 && (result.Value<int>("count") == 0 || result.Value<bool>("visible")), "model metadata fits when provided");
                            Check(result.Value<bool>("metadataMatches"), "picker shows every available multiplier without inventing missing values");
                            Check(result.Value<bool>("fits") && !result.Value<bool>("overflow"), "picker fits narrow panel");
                            await CaptureAsync(core, captureDirectory, $"workbuddy-{width}-picker.png");
                        }
                        else
                        {
                            Check(await core.ExecuteScriptAsync("document.querySelectorAll('.picker-item-multiplier').length === 0") == "true",
                                "unauthorized state does not fabricate model multipliers");
                        }

                        pane.NavigateTo("settings");
                        for (var i = 0; i < 80; i++)
                        {
                            if (await core.ExecuteScriptAsync("!!document.getElementById('workbuddy-login') && document.querySelector('.model-fetch-button')?.disabled === false") == "true") { break; }
                            await Task.Delay(250);
                        }
                        if (authorizationStatus == "\"authorized\"")
                        {
                            Check(await core.ExecuteScriptAsync("(()=>{const options=[...document.getElementById('model-list').options]; return options.length>0 && (options.every(o=>!o.textContent || /\\S/.test(o.textContent)));})()") == "true", "settings model labels contain available models");
                            Check(await core.ExecuteScriptAsync("window.__wbSettings.authorization.models.every(m => !m.multiplier || [...document.getElementById('model-list').options].some(o => o.value===m.modelId && o.textContent.includes(m.multiplier)))") == "true",
                                "settings shows ACP multipliers while retaining model IDs");
                        }
                        else
                        {
                            Check(await core.ExecuteScriptAsync("!!document.getElementById('workbuddy-login') && !document.getElementById('workbuddy-login').disabled") == "true",
                                "unauthorized state keeps browser login control enabled");
                        }
                        Check(await core.ExecuteScriptAsync("document.documentElement.scrollWidth <= innerWidth && [...document.querySelectorAll('.row .btn')].every(n=>{const r=n.getBoundingClientRect();return r.width>0 && r.right<=innerWidth && getComputedStyle(n).whiteSpace==='nowrap';})") == "true", "settings buttons fit and do not wrap");
                        Check(await core.ExecuteScriptAsync(@"(()=>{
                            const buttons=[...document.querySelectorAll('.workbuddy-actions .btn, .model-entry-row .btn')];
                            if (!buttons.length) return false;
                            const rects=buttons.map(n=>n.getBoundingClientRect());
                            return rects.every(r=>r.width>0 && r.height>=24 && r.height<=40 && r.right<=innerWidth+1) &&
                                [...document.querySelectorAll('.workbuddy-action-primary, .workbuddy-action-secondary')]
                                  .every(n=>n.getBoundingClientRect().height<=40);
                        })()") == "true", "WorkBuddy action groups keep stable button heights");
                        Check(await core.ExecuteScriptAsync("(()=>{const b=getComputedStyle(document.body), l=document.querySelector('.field-label'), s=document.querySelector('.workbuddy-status'); return /Segoe UI/i.test(b.fontFamily) && b.fontSize==='13px' && b.lineHeight==='19px' && l && parseFloat(getComputedStyle(l).fontSize)>=13 && s && parseFloat(getComputedStyle(s).fontSize)>=12;})()") == "true", "panel typography uses readable sizes and a crisp UI font stack");
                        if (connectionMode == "\"AuthorizedInternational\"")
                        {
                            Check(await core.ExecuteScriptAsync("document.querySelectorAll('.ready-banner').length===0 && document.querySelectorAll('.field-hint').length===0 && document.querySelectorAll('.workbuddy-status').length===1 && !document.getElementById('settings-status').textContent.trim()") == "true",
                                "international settings use compact status card without stacked explanations");
                            Check(await core.ExecuteScriptAsync("!!document.getElementById('workbuddy-login') && !document.getElementById('workbuddy-login').disabled && (()=>{const n=document.querySelector('.workbuddy-checkin-status'); return n?.textContent==='签到：不适用' && /国内每日签到/.test(n?.title||'');})()") == "true",
                                "international login and no-domestic-checkin status visible");
                        }
                        else
                        {
                            Check(await core.ExecuteScriptAsync("!!document.getElementById('workbuddy-login') && !document.getElementById('workbuddy-login').disabled && !!document.getElementById('workbuddy-checkin-detail')") == "true",
                                "domestic browser login remains actionable and daily checkin controls visible");
                        }
                        Check(await core.ExecuteScriptAsync("document.querySelector('.settings .ready-banner') === null") == "true",
                            "WorkBuddy settings use one status area");
                        Check(await core.ExecuteScriptAsync("document.querySelectorAll('.settings .field-hint').length === 0") == "true",
                            "settings explanations are not permanently stacked");
                        Check(await CompactLayoutAsync(core), "account actions beside status and compact checkin metadata");
                        await CaptureAsync(core, captureDirectory, $"workbuddy-{width}-settings.png");
                        if (connectionMode == "\"AuthorizedInternational\"" && authorizationStatus == "\"authorized\"")
                        {
                            // 只拦截重新登录请求，避免测试登出用户；目录仍使用真实 ACP。
                            await core.ExecuteScriptAsync(@"(() => {
                              const channel=window.chrome.webview, original=channel.postMessage.bind(channel);
                              window.__wbLoginRequests=[];
                              window.__wbSaveRequests=0;
                              window.__wbSelections={};
                              channel.postMessage=m => {
                                if (m.channel==='settings.save') { window.__wbSaveRequests++; }
                                if (m.channel==='workbuddy.login') {
                                  window.__wbLoginRequests.push(m.payload);
                                  queueMicrotask(()=>channel.dispatchEvent(new MessageEvent('message', {data:{
                                    kind:'response',id:m.id,ok:true,data:{ok:true,authorization:window.__wbSettings.authorization}
                                  }})));
                                } else { original(m); }
                              };
                              const change=value=>{const n=document.getElementById('mode');n.value=value;n.dispatchEvent(new Event('change'));};
                              window.__wbChangeMode=change;
                              change('LocalCli');change('AuthorizedInternational');
                            })()");
                            for (var i = 0; i < 120; i++)
                            {
                                if (await core.ExecuteScriptAsync("document.getElementById('workbuddy-login')?.textContent === '切换账号'") == "true") { break; }
                                await Task.Delay(250);
                            }
                            Check(await core.ExecuteScriptAsync("document.getElementById('workbuddy-login')?.textContent === '切换账号' && window.__wbLoginRequests.length===0 && !document.getElementById('model-list').hidden") == "true",
                                "switching back to international restores local authorization and models without login");
                            foreach (var mode in new[] { "Authorized", "AuthorizedInternational" })
                            {
                                await core.ExecuteScriptAsync($"window.__wbChangeMode('{mode}')");
                                for (var i = 0; i < 240; i++)
                                {
                                    if (await core.ExecuteScriptAsync("document.querySelector('.model-fetch-button')?.disabled === false") == "true") { break; }
                                    await Task.Delay(250);
                                }
                                Check(await core.ExecuteScriptAsync("document.querySelector('.workbuddy-status')?.textContent.startsWith('已授权') && !document.getElementById('model-list').hidden && !document.querySelector('.model-fetch-button').disabled") == "true",
                                    mode + " automatically restores official authorization and models");
                                Check(await CompactLayoutAsync(core), mode + " compact layout fits viewport");
                                await core.ExecuteScriptAsync(@"(() => {
                                  const list=document.getElementById('model-list');
                                  list.value=list.options[list.options.length-1]?.value || '';
                                  list.dispatchEvent(new Event('change'));
                                  window.__wbSelections[document.getElementById('mode').value]=document.getElementById('model').value;
                                })()");
                                await CaptureAsync(core, captureDirectory, $"workbuddy-{width}-{mode}-settings.png");
                            }
                            Check(await core.ExecuteScriptAsync(@"(() => {
                              for (const mode of ['Authorized','AuthorizedInternational']) {
                                window.__wbChangeMode(mode);
                                if (document.getElementById('model').value !== window.__wbSelections[mode] ||
                                    document.getElementById('model-list').hidden) return false;
                              }
                              return window.__wbLoginRequests.length===0 && window.__wbSaveRequests===0;
                            })()") == "true", "both editions immediately restore their own selection without login or saving");
                            await core.ExecuteScriptAsync("document.getElementById('workbuddy-login').click()");
                            for (var i = 0; i < 20; i++)
                            {
                                if (await core.ExecuteScriptAsync("!document.getElementById('workbuddy-login').disabled") == "true") { break; }
                                await Task.Delay(100);
                            }
                            Check(await core.ExecuteScriptAsync("window.__wbLoginRequests[0]?.switchAccount === true") == "true",
                                "explicit account switch requests official client account management");
                            await core.ExecuteScriptAsync("window.__wbChangeMode('LocalCli');window.__wbChangeMode('AuthorizedInternational');document.getElementById('workbuddy-login').click()");
                            Check(await core.ExecuteScriptAsync("window.__wbLoginRequests[1]?.switchAccount === true") == "true",
                                "cached authorization retains official account management action");
                        }
                    }
                    catch (Exception ex) { Check(false, ex.ToString()); }
                    finally { form.Close(); }
                };
                Application.Run(form);
            }
            Console.WriteLine($"WorkBuddy UI failures: {failed}");
            return failed == 0 ? 0 : 1;
        }

        private static async Task<bool> CompactLayoutAsync(CoreWebView2 core)
        {
            return await core.ExecuteScriptAsync(@"(() => {
              const row=document.querySelector('.workbuddy-account-row');
              const state=document.querySelector('.workbuddy-state');
              const actions=document.querySelector('.workbuddy-actions');
              const checkin=document.querySelector('.workbuddy-checkin-status');
              if (!row || !state || !actions || !checkin) return false;
              const r=row.getBoundingClientRect(), s=state.getBoundingClientRect(), a=actions.getBoundingClientRect();
              return innerWidth>=360 && document.documentElement.scrollWidth<=innerWidth &&
                a.left>=s.right-1 && a.top<s.bottom && a.bottom>s.top &&
                checkin.getBoundingClientRect().width<r.width/2 &&
                [...actions.querySelectorAll('.btn')].every(n=>{
                  const b=n.getBoundingClientRect();
                  return b.width>0 && b.width<r.width/2 && b.height>=24 && b.height<=40 &&
                    b.right<=innerWidth && getComputedStyle(n).whiteSpace==='nowrap';
                });
            })()") == "true";
        }

        private static async Task CaptureAsync(CoreWebView2 core, string directory, string name)
        {
            if (string.IsNullOrWhiteSpace(directory)) { return; }
            Directory.CreateDirectory(directory);
            using (var stream = File.Create(Path.Combine(directory, name)))
            {
                await core.CapturePreviewAsync(CoreWebView2CapturePreviewImageFormat.Png, stream);
            }
        }
    }
}
