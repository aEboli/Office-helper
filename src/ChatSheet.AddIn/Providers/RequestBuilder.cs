using System;
using System.Collections.Generic;
using ChatSheet.AddIn.Tools;
using Newtonsoft.Json.Linq;

namespace ChatSheet.AddIn.Providers
{
    /// <summary>
    /// 按协议构造请求体。四种协议的消息结构、工具声明与思考参数各不相同，
    /// 差异集中在这里，其余代码只面对统一的 ChatRequest。
    /// </summary>
    internal static class RequestBuilder
    {
        internal static JObject Build(ChatRequest request, bool stream)
        {
            // 出口处统一配平工具调用与结果：配不上的半边会被服务端整轮拒绝，
            // 而它一旦留在历史里就会被每一轮重发。详见 MessageRepair。
            var repaired = MessageRepair.Repair(request.Messages);
            if (repaired.Changed)
            {
                Log.Warn($"请求结构修复：摘掉 {repaired.DroppedCalls} 个没有结果的工具调用、" +
                    $"{repaired.DroppedResults} 条没有调用的工具结果");
            }

            IReadOnlyList<ChatMessage> messages = repaired.Messages;

            // 原生形态聊过几步后切到文本协议（能力回退或用户手动切），历史里仍留着
            // assistant.tool_calls 与 role:tool 的消息，而这次请求体已经不声明 tools。
            // 严格网关对「有工具消息却没声明 tools」直接回 400/502，整轮失败。
            // 降级只改要发出去的那份副本，历史原对象不动：面板的操作卡片与撤销按钮
            // 还按原对象登记着，就地改会让记录与界面各说一套。
            if (!request.IncludeTools)
            {
                messages = DowngradeToolHistory(messages, out var toolCalls, out var toolResults);
                if (toolCalls > 0 || toolResults > 0)
                {
                    Log.Warn($"未声明工具，历史降级：{toolCalls} 条助手调用与 {toolResults} 条工具结果已改为文本");
                }
            }

            switch (request.Protocol)
            {
                case ProtocolKind.AnthropicMessages:
                    return BuildAnthropic(request, messages, stream);
                case ProtocolKind.GoogleGemini:
                    return BuildGemini(request, messages);
                case ProtocolKind.OpenAiResponses:
                    return BuildOpenAiResponses(request, messages, stream);
                default:
                    return BuildOpenAiChat(request, messages, stream);
            }
        }

        /// <summary>
        /// 把历史里的原生工具痕迹降级成纯文本：助手消息的调用并进正文，
        /// 工具结果改写成一条普通 user 消息。
        ///
        /// 为什么需要这一层：IncludeTools 为假表示这次请求体里没有 tools 声明。
        /// 而原生形态聊过的历史里留着 assistant.tool_calls 与 role:"tool" 消息，
        /// 严格网关见到「有工具消息却没声明 tools」会整轮拒绝（400），
        /// 或被当成调用找不到结果的 502。能力回退、用户手动切协议都会走到这里，
        /// 而历史只增不删，不降级就等于每轮都带着它重发。
        ///
        /// 为什么不直接把工具消息删掉：模型会失去「自己做过什么、拿到了什么」，
        /// 下一步必然重复同一批读取。降级保留了信息，只是换了承载形式。
        ///
        /// 为什么不就地改历史：面板上的操作卡片与撤销按钮都还按原对象登记着，
        /// 改掉它等于让记录和界面各说一套。这里只产出复制件。
        ///
        /// 正文里的措辞刻意避开 tool_calls / function_call / tool_use 这些字面量：
        /// 降级后的正文会原样进请求体，写上这些词就等于把要躲开的痕迹又带回去了。
        /// </summary>
        private static IReadOnlyList<ChatMessage> DowngradeToolHistory(
            IReadOnlyList<ChatMessage> source,
            out int downgradedCalls,
            out int downgradedResults)
        {
            downgradedCalls = 0;
            downgradedResults = 0;

            var needed = false;
            foreach (var message in source)
            {
                if (message == null) { continue; }
                if (message.Role == ChatRole.Tool
                    || (message.Role == ChatRole.Assistant && message.ToolCalls.Count > 0))
                {
                    needed = true;
                    break;
                }
            }

            // 没有痕迹就原样交出：多复制一遍会让上下文压缩看到的对象与历史不一致。
            if (!needed)
            {
                return source;
            }

            var downgraded = new List<ChatMessage>(source.Count);
            foreach (var message in source)
            {
                if (message == null)
                {
                    continue;
                }

                if (message.Role == ChatRole.Tool)
                {
                    downgradedResults++;
                    downgraded.Add(new ChatMessage
                    {
                        Role = ChatRole.User,
                        Content = DescribeToolResult(message),
                        // 仍标出工具结果身份：上下文压缩靠它优先压这一类，
                        // 丢了这个标记，压缩会转去丢真正的对话历史。
                        IsTextProtocolToolResult = true,
                    });
                    continue;
                }

                if (message.Role == ChatRole.Assistant && message.ToolCalls.Count > 0)
                {
                    downgradedCalls += message.ToolCalls.Count;
                    var copy = new ChatMessage
                    {
                        Role = ChatRole.Assistant,
                        Content = DescribeAssistantCalls(message),
                    };
                    copy.Images.AddRange(message.Images);
                    downgraded.Add(copy);
                    continue;
                }

                downgraded.Add(message);
            }

            return downgraded;
        }

        /// <summary>把一条助手消息连同它的调用写成正文。调用参数保留，模型才知道当时传了什么。</summary>
        private static string DescribeAssistantCalls(ChatMessage message)
        {
            var builder = new System.Text.StringBuilder();
            if (!string.IsNullOrWhiteSpace(message.Content))
            {
                builder.Append(message.Content.TrimEnd());
                builder.Append('\n');
            }

            builder.Append("（以下为此前的原生调用记录，已改为文本，供了解上下文。）");
            foreach (var call in message.ToolCalls)
            {
                builder.Append("\n工具：").Append(call.Name ?? "(未命名)");
                builder.Append("\n参数：").Append(Trim(call.ArgumentsJson, 600) ?? "{}");
            }

            return builder.ToString();
        }

        /// <summary>把一条工具结果写成普通用户文本。</summary>
        private static string DescribeToolResult(ChatMessage message)
        {
            var name = string.IsNullOrWhiteSpace(message.ToolName) ? "工具" : message.ToolName;
            return $"（工具执行结果 · {name}）\n{message.Content ?? string.Empty}";
        }

        private static string Trim(string text, int limit)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= limit)
            {
                return text;
            }

            return text.Substring(0, limit) + "…（已截断）";
        }

        // ---- OpenAI Chat Completions ----

        private static JObject BuildOpenAiChat(
            ChatRequest request,
            IReadOnlyList<ChatMessage> source,
            bool stream)
        {
            var messages = new JArray();
            foreach (var message in source)
            {
                switch (message.Role)
                {
                    case ChatRole.Tool:
                        messages.Add(new JObject
                        {
                            ["role"] = "tool",
                            ["tool_call_id"] = message.ToolCallId,
                            ["content"] = message.Content ?? string.Empty,
                        });
                        break;

                    case ChatRole.Assistant when message.ToolCalls.Count > 0:
                        var calls = new JArray();
                        foreach (var call in message.ToolCalls)
                        {
                            calls.Add(new JObject
                            {
                                ["id"] = call.Id,
                                ["type"] = "function",
                                ["function"] = new JObject
                                {
                                    ["name"] = call.Name,
                                    ["arguments"] = call.ArgumentsJson ?? "{}",
                                },
                            });
                        }

                        messages.Add(new JObject
                        {
                            ["role"] = "assistant",
                            // 有工具调用时 content 可为空，但字段必须存在。
                            ["content"] = message.Content ?? string.Empty,
                            ["tool_calls"] = calls,
                        });
                        break;

                    default:
                        messages.Add(new JObject
                        {
                            ["role"] = RoleName(message.Role),
                            // 带图片时 content 必须是数组形式。
                            ["content"] = message.Images.Count > 0
                                ? (JToken)BuildOpenAiChatContent(message)
                                : message.Content ?? string.Empty,
                        });
                        break;
                }
            }

            var body = new JObject
            {
                ["model"] = request.Model,
                ["messages"] = messages,
                ["stream"] = stream,
            };

            if (stream)
            {
                // 请求用量统计：默认流式响应不含 usage。
                body["stream_options"] = new JObject { ["include_usage"] = true };
            }

            if (request.Temperature.HasValue) { body["temperature"] = request.Temperature.Value; }
            if (request.MaxOutputTokens.HasValue)
            {
                // 字段名按档案选，不按模型名猜：推理模型只接受 max_completion_tokens，
                // 对 max_tokens 直接回 400。用哪个由服务端拒绝过什么决定。
                var field = request.OutputLimitOverride == OutputLimitField.MaxCompletionTokens
                    ? "max_completion_tokens"
                    : "max_tokens";
                body[field] = request.MaxOutputTokens.Value;
            }

            if (request.IncludeTools)
            {
                body["tools"] = OpenAiTools(request);
                body["tool_choice"] = "auto";
            }

            // OpenAI 系用 reasoning_effort 表达思考强度，
            // 取值 none/minimal/low/medium/high/xhigh/max，各模型支持其中的子集。
            //
            // SuppressThinking 时整段不写：Off 对应的 "none" 仍是一个值，
            // 只认 low/medium/high 的网关会拒绝它。
            if (!request.SuppressThinking)
            {
                var effort = Thinking.OpenAiEffort(request.Thinking);
                if (effort != null)
                {
                    body["reasoning_effort"] = effort;
                }
            }

            return body;
        }

        /// <summary>
        /// Chat Completions 的多模态 content 数组。
        /// 图片用 image_url，其 url 字段接受 data URL 形式的 base64。
        /// </summary>
        private static JArray BuildOpenAiChatContent(ChatMessage message)
        {
            var content = new JArray();

            // 图片放在文本之前：多数视觉模型对「先图后问」的效果更好。
            foreach (var image in message.Images)
            {
                content.Add(new JObject
                {
                    ["type"] = "image_url",
                    ["image_url"] = new JObject { ["url"] = image.ToDataUrl() },
                });
            }

            if (!string.IsNullOrEmpty(message.Content))
            {
                content.Add(new JObject { ["type"] = "text", ["text"] = message.Content });
            }

            return content;
        }

        private static JArray OpenAiTools(ChatRequest request)
        {
            var tools = new JArray();
            foreach (var tool in ToolDefinitions(request))
            {
                tools.Add(new JObject
                {
                    ["type"] = "function",
                    ["function"] = new JObject
                    {
                        ["name"] = tool.Name,
                        ["description"] = tool.Description,
                        ["parameters"] = JObject.FromObject(tool.Parameters),
                    },
                });
            }

            return tools;
        }

        // ---- OpenAI Responses ----

        private static JObject BuildOpenAiResponses(
            ChatRequest request,
            IReadOnlyList<ChatMessage> source,
            bool stream)
        {
            var input = new JArray();
            foreach (var message in source)
            {
                if (message.Role == ChatRole.Tool)
                {
                    input.Add(new JObject
                    {
                        ["type"] = "function_call_output",
                        ["call_id"] = message.ToolCallId,
                        ["output"] = message.Content ?? string.Empty,
                    });
                    continue;
                }

                if (message.Role == ChatRole.Assistant && message.ToolCalls.Count > 0)
                {
                    foreach (var call in message.ToolCalls)
                    {
                        input.Add(new JObject
                        {
                            ["type"] = "function_call",
                            ["call_id"] = call.Id,
                            ["name"] = call.Name,
                            ["arguments"] = call.ArgumentsJson ?? "{}",
                        });
                    }

                    continue;
                }

                input.Add(new JObject
                {
                    ["role"] = RoleName(message.Role),
                    // Responses 协议的图片类型是 input_image，与 Chat Completions 不同。
                    ["content"] = message.Images.Count > 0
                        ? (JToken)BuildResponsesContent(message)
                        : message.Content ?? string.Empty,
                });
            }

            var body = new JObject
            {
                ["model"] = request.Model,
                ["input"] = input,
                ["stream"] = stream,
            };

            if (request.Temperature.HasValue) { body["temperature"] = request.Temperature.Value; }
            if (request.MaxOutputTokens.HasValue) { body["max_output_tokens"] = request.MaxOutputTokens.Value; }

            if (request.IncludeTools)
            {
                var tools = new JArray();
                foreach (var tool in ToolDefinitions(request))
                {
                    // Responses 协议的工具是平铺结构，不像 Chat Completions 那样嵌 function。
                    tools.Add(new JObject
                    {
                        ["type"] = "function",
                        ["name"] = tool.Name,
                        ["description"] = tool.Description,
                        ["parameters"] = JObject.FromObject(tool.Parameters),
                    });
                }

                body["tools"] = tools;
            }

            // SuppressThinking 时整段不写，理由同 BuildOpenAiChat。
            if (!request.SuppressThinking)
            {
                var effort = Thinking.OpenAiEffort(request.Thinking);
                if (effort != null)
                {
                    body["reasoning"] = new JObject { ["effort"] = effort };
                }
            }

            return body;
        }

        /// <summary>
        /// Responses 协议的多模态 content。
        /// 类型名与 Chat Completions 不同：input_text / input_image，
        /// 且 image_url 直接是字符串而非对象。
        /// </summary>
        private static JArray BuildResponsesContent(ChatMessage message)
        {
            var content = new JArray();

            foreach (var image in message.Images)
            {
                content.Add(new JObject
                {
                    ["type"] = "input_image",
                    ["image_url"] = image.ToDataUrl(),
                });
            }

            if (!string.IsNullOrEmpty(message.Content))
            {
                content.Add(new JObject { ["type"] = "input_text", ["text"] = message.Content });
            }

            return content;
        }

        // ---- Anthropic Messages ----

        private static JObject BuildAnthropic(
            ChatRequest request,
            IReadOnlyList<ChatMessage> source,
            bool stream)
        {
            var messages = new JArray();
            string systemPrompt = null;

            foreach (var message in source)
            {
                if (message.Role == ChatRole.System)
                {
                    // Anthropic 的系统提示是顶层字段，不在 messages 里。
                    systemPrompt = string.IsNullOrEmpty(systemPrompt)
                        ? message.Content
                        : systemPrompt + "\n\n" + message.Content;
                    continue;
                }

                if (message.Role == ChatRole.Tool)
                {
                    messages.Add(new JObject
                    {
                        ["role"] = "user",
                        ["content"] = new JArray
                        {
                            new JObject
                            {
                                ["type"] = "tool_result",
                                ["tool_use_id"] = message.ToolCallId,
                                ["content"] = message.Content ?? string.Empty,
                            },
                        },
                    });
                    continue;
                }

                if (message.Role == ChatRole.Assistant && message.ToolCalls.Count > 0)
                {
                    var blocks = new JArray();
                    if (!string.IsNullOrEmpty(message.Content))
                    {
                        blocks.Add(new JObject { ["type"] = "text", ["text"] = message.Content });
                    }

                    foreach (var call in message.ToolCalls)
                    {
                        blocks.Add(new JObject
                        {
                            ["type"] = "tool_use",
                            ["id"] = call.Id,
                            ["name"] = call.Name,
                            ["input"] = ParseArgumentsOrEmpty(call.ArgumentsJson),
                        });
                    }

                    messages.Add(new JObject { ["role"] = "assistant", ["content"] = blocks });
                    continue;
                }

                messages.Add(new JObject
                {
                    ["role"] = RoleName(message.Role),
                    // 带图片时用内容块数组；图片块的 source 是 base64 结构，
                    // 与 OpenAI 的 data URL 形式不同。
                    ["content"] = message.Images.Count > 0
                        ? (JToken)BuildAnthropicContent(message)
                        : message.Content ?? string.Empty,
                });
            }

            var body = new JObject
            {
                ["model"] = request.Model,
                ["messages"] = messages,
                ["stream"] = stream,
                // Anthropic 要求必须显式给出 max_tokens。
                ["max_tokens"] = request.MaxOutputTokens ?? 8192,
            };

            if (!string.IsNullOrEmpty(systemPrompt))
            {
                body["system"] = systemPrompt;
            }

            if (request.IncludeTools)
            {
                var tools = new JArray();
                foreach (var tool in ToolDefinitions(request))
                {
                    tools.Add(new JObject
                    {
                        ["name"] = tool.Name,
                        ["description"] = tool.Description,
                        ["input_schema"] = JObject.FromObject(tool.Parameters),
                    });
                }

                body["tools"] = tools;
            }

            ApplyAnthropicThinking(body, request);
            return body;
        }

        /// <summary>
        /// Anthropic 的多模态内容块。
        /// 图片用 source.type=base64 加 media_type，不接受 data URL。
        /// 官方建议图片置于文字之前，此处照此排列。
        /// </summary>
        private static JArray BuildAnthropicContent(ChatMessage message)
        {
            var content = new JArray();

            foreach (var image in message.Images)
            {
                content.Add(new JObject
                {
                    ["type"] = "image",
                    ["source"] = new JObject
                    {
                        ["type"] = "base64",
                        ["media_type"] = image.MediaType,
                        ["data"] = image.Base64,
                    },
                });
            }

            if (!string.IsNullOrEmpty(message.Content))
            {
                content.Add(new JObject { ["type"] = "text", ["text"] = message.Content });
            }

            return content;
        }

        /// <summary>
        /// 施加 Anthropic 的思考参数。两代模型互不兼容，必须按模型代际二选一：
        ///
        /// - 新方式（4.6 及更新）：thinking.type=adaptive + output_config.effort。
        ///   4.7 及更新的模型只接受这种方式。
        /// - 旧方式（4.5 及更早）：thinking.type=enabled + budget_tokens。
        ///   在 4.7+ 上会被服务端以 400 拒绝。
        ///
        /// 判断依据是模型名（请求前唯一可得的信息）。判断错不致命：
        /// ChatClient 收到相关 400 时会自动改用另一种方式重试。
        /// </summary>
        private static void ApplyAnthropicThinking(JObject body, ChatRequest request)
        {
            // 整段不写。Off 会写 thinking.type="disabled"，那仍是一个值，
            // 而本方法还会顺带写 temperature——探测不该带任何一项。
            if (request.SuppressThinking)
            {
                if (request.Temperature.HasValue)
                {
                    body["temperature"] = request.Temperature.Value;
                }

                return;
            }

            var style = request.AnthropicStyleOverride ?? Thinking.StyleFor(request.Model);

            if (style == AnthropicThinkingStyle.Adaptive)
            {
                if (request.Thinking == ThinkingLevel.Off)
                {
                    body["thinking"] = new JObject { ["type"] = "disabled" };
                }
                else
                {
                    body["thinking"] = new JObject { ["type"] = "adaptive" };
                }

                var effort = Thinking.AnthropicEffort(request.Thinking);
                if (effort != null)
                {
                    body["output_config"] = new JObject { ["effort"] = effort };
                }

                // adaptive 模式不限制 temperature。
                if (request.Temperature.HasValue)
                {
                    body["temperature"] = request.Temperature.Value;
                }

                return;
            }

            var budget = Thinking.AnthropicBudget(request.Thinking, request.MaxOutputTokens ?? 8192);
            if (budget.HasValue)
            {
                body["thinking"] = new JObject
                {
                    ["type"] = "enabled",
                    ["budget_tokens"] = budget.Value,
                };

                // 旧方式启用思考时 temperature 必须为 1，显式传值会被拒绝。
                body.Remove("temperature");
            }
            else if (request.Temperature.HasValue)
            {
                body["temperature"] = request.Temperature.Value;
            }
        }

        // ---- Google Gemini ----

        private static JObject BuildGemini(ChatRequest request, IReadOnlyList<ChatMessage> source)
        {
            var contents = new JArray();
            JObject systemInstruction = null;

            foreach (var message in source)
            {
                if (message.Role == ChatRole.System)
                {
                    systemInstruction = new JObject
                    {
                        ["parts"] = new JArray { new JObject { ["text"] = message.Content ?? string.Empty } },
                    };
                    continue;
                }

                if (message.Role == ChatRole.Tool)
                {
                    contents.Add(new JObject
                    {
                        ["role"] = "user",
                        ["parts"] = new JArray
                        {
                            new JObject
                            {
                                ["functionResponse"] = new JObject
                                {
                                    ["name"] = message.ToolName ?? string.Empty,
                                    // Gemini 要求响应是对象，字符串需包一层。
                                    ["response"] = new JObject { ["result"] = message.Content ?? string.Empty },
                                },
                            },
                        },
                    });
                    continue;
                }

                if (message.Role == ChatRole.Assistant && message.ToolCalls.Count > 0)
                {
                    var callParts = new JArray();
                    if (!string.IsNullOrEmpty(message.Content))
                    {
                        callParts.Add(new JObject { ["text"] = message.Content });
                    }

                    foreach (var call in message.ToolCalls)
                    {
                        callParts.Add(new JObject
                        {
                            ["functionCall"] = new JObject
                            {
                                ["name"] = call.Name,
                                ["args"] = ParseArgumentsOrEmpty(call.ArgumentsJson),
                            },
                        });
                    }

                    contents.Add(new JObject { ["role"] = "model", ["parts"] = callParts });
                    continue;
                }

                var parts = new JArray();

                // Gemini 的图片用 inlineData，字段名为 mimeType（驼峰）。
                foreach (var image in message.Images)
                {
                    parts.Add(new JObject
                    {
                        ["inlineData"] = new JObject
                        {
                            ["mimeType"] = image.MediaType,
                            ["data"] = image.Base64,
                        },
                    });
                }

                parts.Add(new JObject { ["text"] = message.Content ?? string.Empty });

                contents.Add(new JObject
                {
                    // Gemini 用 model 而非 assistant。
                    ["role"] = message.Role == ChatRole.Assistant ? "model" : "user",
                    ["parts"] = parts,
                });
            }

            var body = new JObject { ["contents"] = contents };

            if (systemInstruction != null)
            {
                body["systemInstruction"] = systemInstruction;
            }

            var generationConfig = new JObject();
            if (request.Temperature.HasValue) { generationConfig["temperature"] = request.Temperature.Value; }
            if (request.MaxOutputTokens.HasValue) { generationConfig["maxOutputTokens"] = request.MaxOutputTokens.Value; }

            // Gemini 的 thinkingConfig 支持两种表达：
            // thinkingLevel（新，minimal/low/medium/high）与 thinkingBudget（旧，token 数）。
            // 这里用 thinkingLevel，并附 includeThoughts 以便展示思考过程。
            // thinkingBudget=0 是官方表达「关闭思考」的方式。
            // SuppressThinking 时整段不写 thinkingConfig。
            // 注意 else 分支写的 thinkingBudget=0 也是一个值——不写才是不写。
            if (!request.SuppressThinking)
            {
                var geminiLevel = Thinking.GeminiLevel(request.Thinking);
                if (geminiLevel != null)
                {
                    generationConfig["thinkingConfig"] = new JObject
                    {
                        ["thinkingLevel"] = geminiLevel,
                        ["includeThoughts"] = true,
                    };
                }
                else
                {
                    generationConfig["thinkingConfig"] = new JObject { ["thinkingBudget"] = 0 };
                }
            }

            if (generationConfig.Count > 0)
            {
                body["generationConfig"] = generationConfig;
            }

            if (request.IncludeTools)
            {
                var declarations = new JArray();
                foreach (var tool in ToolDefinitions(request))
                {
                    declarations.Add(new JObject
                    {
                        ["name"] = tool.Name,
                        ["description"] = tool.Description,
                        ["parameters"] = SanitizeForGemini(JObject.FromObject(tool.Parameters)),
                    });
                }

                body["tools"] = new JArray { new JObject { ["functionDeclarations"] = declarations } };
            }

            return body;
        }

        /// <summary>
        /// Gemini 的 schema 子集不接受 additionalProperties，
        /// 且不支持 type 为数组的联合类型，需要就近降级为单一类型。
        /// </summary>
        private static JToken SanitizeForGemini(JToken token)
        {
            if (token is JObject obj)
            {
                var result = new JObject();
                foreach (var property in obj.Properties())
                {
                    if (string.Equals(property.Name, "additionalProperties", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    if (string.Equals(property.Name, "type", StringComparison.Ordinal) &&
                        property.Value is JArray typeArray)
                    {
                        // 取第一个非 null 类型作为代表。
                        string chosen = null;
                        foreach (var item in typeArray)
                        {
                            var name = item.Value<string>();
                            if (!string.Equals(name, "null", StringComparison.OrdinalIgnoreCase))
                            {
                                chosen = name;
                                break;
                            }
                        }

                        result["type"] = chosen ?? "string";
                        continue;
                    }

                    result[property.Name] = SanitizeForGemini(property.Value);
                }

                return result;
            }

            if (token is JArray array)
            {
                var result = new JArray();
                foreach (var item in array)
                {
                    result.Add(SanitizeForGemini(item));
                }

                return result;
            }

            return token;
        }

        // ---- 公共辅助 ----

        private static string RoleName(ChatRole role)
        {
            switch (role)
            {
                case ChatRole.System: return "system";
                case ChatRole.Assistant: return "assistant";
                case ChatRole.Tool: return "tool";
                default: return "user";
            }
        }

        private static IReadOnlyList<ToolDefinition> ToolDefinitions(ChatRequest request)
        {
            return request.ToolDefinitions ?? ToolCatalog.All;
        }

        private static JObject ParseArgumentsOrEmpty(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return new JObject();
            }

            try
            {
                return JObject.Parse(json);
            }
            catch
            {
                // 模型偶尔会产出不完整的 JSON，此时退化为空参数，
                // 让工具层给出「缺少参数」的明确错误，而不是在此崩溃。
                return new JObject();
            }
        }
    }
}
