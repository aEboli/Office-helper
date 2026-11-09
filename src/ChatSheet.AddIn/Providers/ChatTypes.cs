using System.Collections.Generic;
using System.Threading.Tasks;
using ChatSheet.AddIn.Tools;

namespace ChatSheet.AddIn.Providers
{
    internal enum ChatRole
    {
        System = 0,
        User = 1,
        Assistant = 2,
        Tool = 3,
    }

    /// <summary>模型请求的工具调用。</summary>
    internal sealed class ToolCall
    {
        internal string Id { get; set; }

        internal string Name { get; set; }

        /// <summary>参数的 JSON 文本。流式场景下按增量拼接而成。</summary>
        internal string ArgumentsJson { get; set; }
    }

    internal sealed class ChatMessage
    {
        internal ChatRole Role { get; set; }

        internal string Content { get; set; }

        /// <summary>随本条消息发送的图片。仅用户消息会带。</summary>
        internal List<ImageAttachment> Images { get; } = new List<ImageAttachment>();

        /// <summary>助手消息附带的工具调用。</summary>
        internal List<ToolCall> ToolCalls { get; } = new List<ToolCall>();

        /// <summary>工具结果消息对应的调用标识。</summary>
        internal string ToolCallId { get; set; }

        /// <summary>工具名称，Gemini 与 Anthropic 的结果消息需要它。</summary>
        internal string ToolName { get; set; }

        /// <summary>
        /// 这是文本协议下以 user 消息回灌的工具结果。
        ///
        /// 需要单独标记：文本协议没有 tool_call_id，结果只能作为普通用户消息发出，
        /// 于是上下文压缩认不出它是工具结果——而工具结果恰是体积最大、
        /// 最该优先压缩的那一类。丢了这个标记，压缩就会转去丢弃真正的对话历史。
        /// </summary>
        internal bool IsTextProtocolToolResult { get; set; }

        internal static ChatMessage FromSystem(string content) =>
            new ChatMessage { Role = ChatRole.System, Content = content };

        internal static ChatMessage FromUser(string content) =>
            new ChatMessage { Role = ChatRole.User, Content = content };

        /// <summary>带图片的用户消息。</summary>
        internal static ChatMessage FromUser(string content, IEnumerable<ImageAttachment> images)
        {
            var message = new ChatMessage { Role = ChatRole.User, Content = content };
            if (images != null)
            {
                message.Images.AddRange(images);
            }

            return message;
        }

        internal static ChatMessage FromAssistant(string content) =>
            new ChatMessage { Role = ChatRole.Assistant, Content = content };

        internal static ChatMessage FromToolResult(string toolCallId, string toolName, string content) =>
            new ChatMessage
            {
                Role = ChatRole.Tool,
                ToolCallId = toolCallId,
                ToolName = toolName,
                Content = content,
            };

        /// <summary>
        /// 文本协议下的工具结果。角色是 user，因为协议里没有工具消息可用，
        /// 但仍标记出身份供上下文压缩识别。
        /// </summary>
        internal static ChatMessage FromTextProtocolToolResult(string toolName, string content) =>
            new ChatMessage
            {
                Role = ChatRole.User,
                ToolName = toolName,
                Content = content,
                IsTextProtocolToolResult = true,
            };

        /// <summary>
        /// 复制一条消息，只替换附带的工具调用，正文可选替换。
        ///
        /// 请求修复要摘掉没有结果的调用，但不能就地改历史：面板上的操作卡片与
        /// 撤销按钮都还按原对象登记着，改掉它等于让记录和界面各说一套。
        /// </summary>
        internal ChatMessage CloneWithToolCalls(IReadOnlyList<ToolCall> calls, string content = null)
        {
            var clone = new ChatMessage
            {
                Role = Role,
                Content = content ?? Content,
                ToolCallId = ToolCallId,
                ToolName = ToolName,
                IsTextProtocolToolResult = IsTextProtocolToolResult,
            };

            if (calls != null)
            {
                clone.ToolCalls.AddRange(calls);
            }

            clone.Images.AddRange(Images);
            return clone;
        }
    }

    /// <summary>Chat Completions 上输出上限的字段名。</summary>
    internal enum OutputLimitField
    {
        /// <summary>历来的写法，绝大多数模型接受。</summary>
        MaxTokens = 0,

        /// <summary>OpenAI 推理模型只接受这个，对 max_tokens 回 400。</summary>
        MaxCompletionTokens = 1,
    }

    internal sealed class ChatRequest
    {
        internal ProtocolKind Protocol { get; set; }

        internal string BaseUrl { get; set; }

        internal string Token { get; set; }

        internal string Model { get; set; }

        internal List<ChatMessage> Messages { get; } = new List<ChatMessage>();

        internal ThinkingLevel Thinking { get; set; } = ThinkingLevel.Off;

        internal double? Temperature { get; set; }

        internal int? MaxOutputTokens { get; set; }

        /// <summary>是否附带工具声明。纯问答场景可关闭以省 token。</summary>
        internal bool IncludeTools { get; set; } = true;

        /// <summary>
        /// 本次请求使用的宿主工具定义。为空时沿用 Excel 工具目录，
        /// 供独立的 Word 加载项注入文档工具清单。
        /// </summary>
        internal IReadOnlyList<ToolDefinition> ToolDefinitions { get; set; }

        /// <summary>
        /// 整段不写思考参数，而不是写一个「关闭」值。
        ///
        /// 两者不同，而这个区别会决定探测的结论对不对：Thinking = Off 实际会发出
        /// reasoning_effort:"none"（OpenAI 系）或 thinkingBudget:0（Gemini），
        /// 那仍然是一个值。只认 low/medium/high 的网关会以 400 拒绝它，
        /// 于是「这个模型能不能用」永远问不出答案；反过来，一个会拒绝真实对话所用
        /// 档位的模型，在 "none" 下可能通过，探测就给出一个骗人的绿灯。
        ///
        /// 置真时请求体在思考这一项上是真实请求的真子集。
        /// </summary>
        internal bool SuppressThinking { get; set; }

        /// <summary>
        /// Chat Completions 上输出上限用哪个字段名。为空表示沿用默认的 max_tokens。
        ///
        /// OpenAI 的推理模型只接受 max_completion_tokens，对 max_tokens 直接回 400。
        /// 用哪个由服务端的拒绝决定，不按模型名猜——模型名与行为没有可靠对应关系。
        /// </summary>
        internal OutputLimitField? OutputLimitOverride { get; set; }

        /// <summary>
        /// 强制指定 Anthropic 的思考控制方式，为空则按模型名推断。
        /// 用于服务端以 400 拒绝某种方式后自动改用另一种重试。
        /// </summary>
        internal AnthropicThinkingStyle? AnthropicStyleOverride { get; set; }
    }

    /// <summary>流式事件类型。</summary>
    internal enum ChatEventKind
    {
        /// <summary>正文增量。</summary>
        TextDelta = 0,

        /// <summary>思考过程增量。</summary>
        ThinkingDelta = 1,

        /// <summary>工具调用已完整解析。</summary>
        ToolCall = 2,

        /// <summary>本轮结束。</summary>
        Completed = 3,

        /// <summary>用量统计。</summary>
        Usage = 4,
    }

    internal sealed class ChatEvent
    {
        internal ChatEventKind Kind { get; set; }

        internal string Text { get; set; }

        internal ToolCall Call { get; set; }

        internal int PromptTokens { get; set; }

        internal int CompletionTokens { get; set; }

        /// <summary>结束原因，例如 stop、tool_calls、length。</summary>
        internal string FinishReason { get; set; }
    }

    /// <summary>把不同接入方式归一成 AgentRunner 使用的流式客户端。</summary>
    internal interface IChatStreamClient : System.IDisposable
    {
        Task StreamAsync(
            ChatRequest request,
            System.Func<ChatEvent, Task> onEvent,
            System.Threading.CancellationToken cancellationToken,
            System.Func<int, System.TimeSpan, string, Task> onRetry = null,
            int? maxRetries = null);
    }
}
