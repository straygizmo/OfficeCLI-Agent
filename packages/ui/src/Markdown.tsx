import { memo } from "react";
import ReactMarkdown, { type Components } from "react-markdown";
import remarkBreaks from "remark-breaks";
import remarkGfm from "remark-gfm";

// GFM for tables/task lists/strikethrough; single line breaks stay line breaks, as in the plain-text view before.
const plugins = [remarkGfm, remarkBreaks];

const components: Components = {
  // Links leave the task pane: the add-in opens new-window requests in the default browser.
  a: ({ node: _node, ...props }) => <a {...props} target="_blank" rel="noreferrer" />,
  // Wide tables scroll inside the bubble instead of stretching the narrow pane.
  table: ({ node: _node, ...props }) => (
    <div className="md-table">
      <table {...props} />
    </div>
  ),
};

/** Assistant reply rendered as Markdown. Raw HTML in the text is not interpreted (shown as text). */
export const Markdown = memo(function Markdown({ text }: { text: string }) {
  return (
    <ReactMarkdown remarkPlugins={plugins} components={components}>
      {text}
    </ReactMarkdown>
  );
});
