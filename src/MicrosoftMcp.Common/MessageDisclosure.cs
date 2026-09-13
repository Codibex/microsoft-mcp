namespace MicrosoftMcp.Common;

/// <summary>Server-side AI-disclosure injection for mail and (future) Teams messages.
/// The text is appended by the server, never supplied by the LLM, so it cannot be
/// omitted or reworded via prompt. Idempotent: re-applying (e.g. via
/// outlook_update_draft) does not duplicate it. Pure and unit-tested.</summary>
public static class MessageDisclosure
{
    /// <summary>Applies the disclosure, auto-detecting HTML vs. text bodies
    /// (reply/forward comments may contain either).</summary>
    public static string ApplyAuto(string body, MessagingPolicyOptions policy) =>
        Apply(body, LooksLikeHtml(body), policy);

    /// <summary>Heuristic: true when the text contains a tag-like construct
    /// (&lt;letter, &lt;/, &lt;! or &lt;? … &gt;). "a &lt; b" stays text.</summary>
    public static bool LooksLikeHtml(string value)
    {
        for (int i = 0; i < value.Length - 1; i++)
        {
            if (value[i] != '<')
            {
                continue;
            }

            char next = value[i + 1];
            if (!char.IsAsciiLetter(next) && next is not ('/' or '!' or '?'))
            {
                continue;
            }

            if (value.IndexOf('>', i + 2) > i + 2)
            {
                return true;
            }
        }

        return false;
    }

    public static string Apply(string body, bool isHtml, MessagingPolicyOptions policy)
    {
        if (!policy.AiDisclosureEnabled || string.IsNullOrWhiteSpace(policy.AiDisclosureText))
        {
            return body;
        }

        string text = policy.AiDisclosureText.Trim();
        if (isHtml)
        {
            string encoded = System.Net.WebUtility.HtmlEncode(text);
            if (body.Contains(encoded, StringComparison.Ordinal) || body.Contains(text, StringComparison.Ordinal))
            {
                return body;
            }

            return body.Length == 0 ? $"<p>{encoded}</p>" : $"{body}<hr><p>{encoded}</p>";
        }

        if (body.Contains(text, StringComparison.Ordinal))
        {
            return body;
        }

        return body.Length == 0 ? text : $"{body}\n\n---\n{text}";
    }
}
