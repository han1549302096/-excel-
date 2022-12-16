import tkinter as tk
import openai

# 设置 API 密钥
openai.api_key = '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'

class TextCompletionApp:
    def __init__(self, window):
        # 创建标签和文本输入框
        self.label = tk.Label(window, text="输入提示：")
        self.entry = tk.Entry(window)

        # 创建一个按钮小部件
        button = tk.Button(window, text="Submit", command=self.submit)

        # 使用网格布局管理器将小部件放在窗口中
        self.label.grid(row=0, column=0)
        self.entry.grid(row=0, column=1)
        button.grid(row=1, column=0, columnspan=2)

        # 创建 Text 组件
        self.output_text = tk.Text(window)
        self.output_text.grid(row=2, column=0, columnspan=2)

    def submit(self):
        # 从文本输入框中获取值
        prompt = self.entry.get()

        # 如果输入提示为空，则不进行任何操作
        if not prompt:
            return

        try:
            # 调用 OpenAI 的 API
            response = openai.Completion.create(
                model='code-cushman-001',
                prompt=prompt,
                temperature=2.0,
                max_tokens=200,
                top_p=1.0,
                frequency_penalty=0.0,
                presence_penalty=0.0,
                stop=['#', '\"\"\"'],
            )

            # 确保响应是有效的
            if response.status == "success" and len(response.choices) > 0:
                # 在 Text 组件中追加新的文本输出
                self.output_text.mark_set("insert", "end")
                self.output_text.insert("end", response.choices[0].text)
            else:
                # 如果响应不是有效的，显示错误消息
                self.output_text.mark_set("insert", "end")
                self.output_text.insert("1.0", "Error: Invalid response from API")
        except Exception as e:
            # 如果发生异常，显示错误信息
            self.output_text.mark_set("insert", "end")
            self.output_text.insert("1.0", f"Error: {e}")

    def clear(self):
        self.output_text.delete("1.0", "end")




# 创建主窗口
window = tk.Tk()
window.title("Text Completion")
# 创建清除按钮
clear_button = tk.Button(window, text="Clear", command=self.clear)
clear_button.grid(row=3, column=0, columnspan=2)
# 创建 TextCompletionApp 的实例
app = TextCompletionApp(window)

# 运行主循环
while True:
    window.update()
    window.mainloop()


#我想将这个代码生成的界面改的精致一些