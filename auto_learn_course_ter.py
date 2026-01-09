import tkinter as tk
from tkinter import ttk
import threading
import time
import logging
import sys
import os
import re
from selenium import webdriver
from selenium.webdriver import ActionChains  # 新增：鼠标模拟点击
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException, \
    JavascriptException
from webdriver_manager.chrome import ChromeDriverManager
from queue import Queue

# ===================== 全局配置 =====================
TARGET_URL = "https://sdld-gxk.yxlearning.com/my/index"
WAIT_TIMEOUT = 30
RETRY_TIMES = 3  # 元素操作失败重试次数
# 全局线程控制
STOP_EVENT = threading.Event()  # 停止进程（退出所有程序）
PAUSE_EVENT = threading.Event()  # 暂停任务（继续任务时清除）
STATUS_QUEUE = Queue()  # 线程安全的状态更新队列
LOG_FILE = "auto_learn_log.log"  # 日志文件
# 全局锁（保护Driver操作）
DRIVER_LOCK = threading.Lock()
driver = None  # 全局Driver对象
is_learn_started = False  # 标记是否已触发学习流程


# ===================== 日志系统初始化 =====================
def init_logger():
    logger = logging.getLogger("AutoLearn")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    # 文件处理器
    file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
    file_formatter = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(threadName)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    file_handler.setFormatter(file_formatter)

    # 控制台处理器
    console_handler = logging.StreamHandler()
    console_formatter = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%H:%M:%S"
    )
    console_handler.setFormatter(console_formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return logger


logger = init_logger()


# ===================== 通用工具函数（核心修复：适配SVG点击） =====================
def safe_click_element(driver, element):
    """安全点击元素：优先JS点击，失败则用鼠标模拟点击（适配SVG元素）"""
    try:
        # 方案1：JS点击（优先）
        driver.execute_script("arguments[0].click();", element)
        update_status("使用JS成功点击元素")
        return True
    except JavascriptException as e:
        update_status(f"JS点击失败（SVG元素）：{str(e)[:30]}，切换为鼠标模拟点击")
        logger.warning(f"JS点击SVG失败：{e}")
        try:
            # 方案2：鼠标模拟点击（适配SVG）
            ActionChains(driver).move_to_element(element).click().perform()
            update_status("使用鼠标模拟点击成功")
            return True
        except Exception as e2:
            update_status(f"鼠标模拟点击也失败：{str(e2)[:30]}")
            logger.warning(f"鼠标点击失败：{e2}")
            return False
    except Exception as e:
        update_status(f"点击元素失败：{str(e)[:30]}")
        logger.warning(f"点击失败：{e}")
        return False


def extract_progress_percent(progress_text):
    """解析进度文本，提取百分比数值（适配：总学时进度（26.47%））"""
    try:
        # 正则匹配括号内的百分比数值（如26.47）
        match = re.search(r'(\d+\.?\d*)%', progress_text)
        if match:
            return float(match.group(1))
        # 兼容纯文本状态（如“您已完成该课程的学习”）
        if "已完成" in progress_text:
            return 100.0
        return 0.0
    except Exception as e:
        logger.warning(f"解析进度失败：{e}，文本：{progress_text}")
        return 0.0


# ===================== 通用重试装饰器 =====================
def retry_on_failure(max_retries=RETRY_TIMES, delay=2):
    def decorator(func):
        def wrapper(*args, **kwargs):
            for retry in range(max_retries):
                if PAUSE_EVENT.is_set() or STOP_EVENT.is_set():
                    raise Exception("任务已暂停或停止")
                try:
                    return func(*args, **kwargs)
                except (TimeoutException, NoSuchElementException, ElementClickInterceptedException) as e:
                    logger.warning(f"【重试{retry + 1}/{max_retries}】{func.__name__}失败：{str(e)[:50]}")
                    time.sleep(delay)
            raise Exception(f"【重试耗尽】{func.__name__}失败，已重试{max_retries}次")

        return wrapper

    return decorator


# ===================== 状态更新工具 =====================
def update_status(msg):
    STATUS_QUEUE.put(msg)
    logger.info(f"【状态更新】{msg}")


# ===================== 元素查找函数 =====================
@retry_on_failure(max_retries=RETRY_TIMES)
def find_element_clickable(driver, locator, timeout=WAIT_TIMEOUT):
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable(locator)
    )


@retry_on_failure(max_retries=RETRY_TIMES)
def find_element_present(driver, locator, timeout=WAIT_TIMEOUT):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located(locator)
    )


# ===================== 学习流程初始化函数 =====================
def init_learn_flow(driver):
    """初始化学习流程：点击「我的学习」→「全部年度」"""
    global is_learn_started
    if is_learn_started:
        update_status("学习流程已初始化，无需重复执行")
        return True

    try:
        update_status("开始初始化学习流程：点击「我的学习」")
        # 定位「我的学习」
        my_learning_locator = (By.XPATH, "//*[contains(text(), '我的学习')]")
        my_learning_clickable = find_element_clickable(driver, my_learning_locator, timeout=10)
        safe_click_element(driver, my_learning_clickable)
        # 等待页面跳转
        for _ in range(5):
            if PAUSE_EVENT.is_set() or STOP_EVENT.is_set():
                return False
            time.sleep(1)

        update_status("学习流程初始化：点击「全部年度」")
        # 定位「全部年度」
        all_year_locator = (By.XPATH, "//div[@class='yearItem bg-white' and .//p[@class='year' and text()='全部年度']]")
        all_year_tab = find_element_clickable(driver, all_year_locator)
        safe_click_element(driver, all_year_tab)
        # 等待标签加载稳定
        for _ in range(5):
            if PAUSE_EVENT.is_set() or STOP_EVENT.is_set():
                return False
            time.sleep(1)

        is_learn_started = True
        update_status("学习流程初始化完成：已进入「全部年度」页面")
        return True
    except Exception as e:
        update_status(f"学习流程初始化失败：{str(e)[:30]}")
        logger.error(f"初始化学习流程出错：{e}", exc_info=True)
        return False


# ===================== 自动化核心任务（最终最终版） =====================
def auto_learn_task():
    """自动化核心任务：打开页面→等待手动登录→执行学习流程→循环学习"""
    global driver
    try:
        update_status("初始化Chrome浏览器")
        # Chrome配置优化
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument("--disable-cache")
        chrome_options.add_argument("--disk-cache-size=0")

        with DRIVER_LOCK:
            driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()),
                options=chrome_options
            )
            driver.implicitly_wait(5)
            driver.set_page_load_timeout(WAIT_TIMEOUT)
            driver.maximize_window()

            # 打开初始页面，暂停等待手动登录
            update_status("打开初始页面：请手动登录账号！（程序不自动处理弹窗）")
            driver.get(TARGET_URL)
            time.sleep(2)

        # 初始状态：暂停任务
        PAUSE_EVENT.set()
        update_status("页面已打开，任务暂停，请手动登录并关闭弹窗")

        cycle_count = 0
        while not STOP_EVENT.is_set():
            # 检测暂停状态
            while PAUSE_EVENT.is_set() and not STOP_EVENT.is_set():
                update_status("等待用户操作：登录完成后点击悬浮窗「开始任务」")
                time.sleep(1)
            if STOP_EVENT.is_set():
                break

            cycle_count += 1
            update_status(f"开始第{cycle_count}轮学习任务")

            # 首次执行学习流程初始化
            with DRIVER_LOCK:
                if driver and not is_learn_started:
                    update_status("首次执行：初始化学习流程（我的学习→全部年度）")
                    if not init_learn_flow(driver):
                        update_status("学习流程初始化失败，暂停任务")
                        PAUSE_EVENT.set()
                        continue
            # ==================================================================================

            if STOP_EVENT.is_set() or PAUSE_EVENT.is_set():
                break

            # ===================== 继续学习按钮（适配你的HTML） =====================
            update_status(f"第{cycle_count}轮：检测并点击「继续学习」按钮")
            # 定位器：文本“继续学习” + class="item-bottom-btn"
            continue_locator = (By.XPATH, "//div[contains(text(), '继续学习') and contains(@class, 'item-bottom-btn')]")
            with DRIVER_LOCK:
                continue_learn_btn = find_element_clickable(driver, continue_locator)
                safe_click_element(driver, continue_learn_btn)
            if PAUSE_EVENT.is_set() or STOP_EVENT.is_set():
                break

            # ===================== 进度监听（适配jindu-span+特殊文本格式） =====================
            update_status(f"第{cycle_count}轮：监听课程完成状态")
            progress_check_count = 0
            is_course_finished = False
            while not STOP_EVENT.is_set() and not is_course_finished:
                if PAUSE_EVENT.is_set():
                    break
                progress_check_count += 1
                if progress_check_count % 3 == 0:
                    update_status(f"第{cycle_count}轮：进度监听中（心跳{progress_check_count}）")

                try:
                    with DRIVER_LOCK:
                        # 匹配你实际的进度类 jindu-span
                        progress_elem = find_element_present(driver, (By.CLASS_NAME, "jindu-span"))
                        progress_text = progress_elem.text.strip()
                    # 提取百分比数值（如从“总学时进度（26.47%）”提取26.47）
                    progress_percent = extract_progress_percent(progress_text)
                    update_status(f"第{cycle_count}轮：课程进度：{progress_text}（提取数值：{progress_percent}%）")

                    # 判断完成条件：数值≥100 或 文本含“已完成”
                    if progress_percent >= 100.0 or "已完成" in progress_text:
                        update_status(f"第{cycle_count}轮：课程学习完成！准备返回")
                        is_course_finished = True
                        break
                    else:
                        # 分段sleep，检测暂停/停止
                        for _ in range(10):
                            if PAUSE_EVENT.is_set() or STOP_EVENT.is_set():
                                break
                            time.sleep(1)
                except Exception as e:
                    update_status(f"第{cycle_count}轮：检查进度失败，10秒后重试")
                    logger.warning(f"进度检查失败：{e}")
                    for _ in range(10):
                        if PAUSE_EVENT.is_set() or STOP_EVENT.is_set():
                            break
                        time.sleep(1)
            if PAUSE_EVENT.is_set() or STOP_EVENT.is_set():
                break
            if not is_course_finished:
                continue  # 未完成则跳过返回步骤

            # ===================== 返回按钮（核心修复：定位SVG父元素+适配点击） =====================
            update_status(f"第{cycle_count}轮：点击返回按钮")
            with DRIVER_LOCK:
                # 修复定位器：找到svg.back-icon的父元素（可点击容器），而非直接定位svg
                back_locator = (By.XPATH, "//svg[@class='back-icon']/parent::*[1]")
                # 备选定位器：如果父元素找不到，直接定位svg
                try:
                    back_btn = find_element_clickable(driver, back_locator, timeout=10)
                except:
                    update_status("未找到返回按钮父元素，直接定位SVG元素")
                    back_btn = find_element_clickable(driver, (By.CLASS_NAME, "back-icon"), timeout=10)
                # 使用安全点击函数（适配SVG）
                safe_click_element(driver, back_btn)
            # 等待返回页面加载稳定
            for _ in range(5):
                if PAUSE_EVENT.is_set() or STOP_EVENT.is_set():
                    break
                time.sleep(1)
            update_status(f"第{cycle_count}轮：返回成功，准备下一轮学习")
        # ==================================================================================

    except Exception as e:
        if not STOP_EVENT.is_set():
            update_status(f"任务出错：{str(e)[:30]}")
            logger.error(f"任务执行出错：{e}", exc_info=True)
    finally:
        # 释放资源
        with DRIVER_LOCK:
            if driver:
                update_status("关闭Chrome浏览器，释放资源")
                driver.quit()
        update_status("任务结束，所有资源已释放")


# ===================== 悬浮窗类（保留所有功能） =====================
class TaskFloatWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("学习任务监控")
        self.root.attributes("-topmost", True)
        self.root.overrideredirect(True)
        self.root.geometry("400x150+100+100")
        self.root.attributes("-alpha", 0.9)

        # 配置按钮样式
        self.style = ttk.Style()
        self.style.configure("Accent.TButton", font=("微软雅黑", 9))
        self.style.configure("Warning.TButton", font=("微软雅黑", 9), foreground="red")

        # 状态标签
        self.status_label = ttk.Label(
            self.root,
            text="当前状态：页面已打开，请手动登录并关闭弹窗，登录完成后点击「开始任务」",
            font=("微软雅黑", 10),
            wraplength=380,
            justify="left"
        )
        self.status_label.pack(pady=10, padx=10)

        # 按钮框架
        self.btn_frame = ttk.Frame(self.root)
        self.btn_frame.pack(pady=5, fill=tk.X, padx=10)

        # 1. 停止任务按钮
        self.pause_btn = ttk.Button(
            self.btn_frame,
            text="停止任务",
            command=self.pause_task,
            style="Accent.TButton",
            width=12
        )
        self.pause_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # 2. 开始任务按钮
        self.resume_btn = ttk.Button(
            self.btn_frame,
            text="开始任务",
            command=self.resume_task,
            style="Accent.TButton",
            width=12
        )
        self.resume_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # 3. 停止进程按钮
        self.stop_btn = ttk.Button(
            self.btn_frame,
            text="停止进程",
            command=self.stop_process,
            style="Warning.TButton",
            width=12
        )
        self.stop_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # 拖动功能
        self.root.bind("<Button-1>", self.start_drag)
        self.root.bind("<B1-Motion>", self.on_drag)
        self.x = 0
        self.y = 0

        # 启动任务线程
        self.task_thread = threading.Thread(target=auto_learn_task, name="AutoLearnThread", daemon=True)
        self.task_thread.start()

        # 消费状态队列
        self.consume_status_queue()

    def start_drag(self, event):
        self.x = event.x
        self.y = event.y

    def on_drag(self, event):
        deltax = event.x - self.x
        deltay = event.y - self.y
        x = self.root.winfo_x() + deltax
        y = self.root.winfo_y() + deltay
        self.root.geometry(f"400x150+{x}+{y}")

    def pause_task(self):
        """暂停当前循环，保留浏览器"""
        if not PAUSE_EVENT.is_set():
            PAUSE_EVENT.set()
            update_status("用户手动暂停任务")
            self.status_label.config(text="当前状态：任务已暂停，点击「开始任务」继续")
        else:
            update_status("任务已处于暂停状态")

    def resume_task(self):
        """继续任务"""
        global driver
        if STOP_EVENT.is_set():
            update_status("进程已停止，无法继续任务")
            self.status_label.config(text="当前状态：进程已停止，无法继续")
            return

        PAUSE_EVENT.clear()
        update_status("用户手动启动任务，开始执行学习流程")
        self.status_label.config(text="当前状态：任务已启动，正在执行学习流程")

    def stop_process(self):
        """停止所有进程，关闭Chrome和悬浮窗"""
        update_status("用户手动停止进程，正在退出所有程序")
        STOP_EVENT.set()
        PAUSE_EVENT.clear()

        # 关闭Chrome
        global driver
        with DRIVER_LOCK:
            if driver:
                try:
                    driver.quit()
                except:
                    pass

        # 关闭悬浮窗
        self.root.quit()
        self.root.destroy()
        logger.info("进程已完全停止，所有资源已释放")
        sys.exit(0)

    def consume_status_queue(self):
        """消费状态队列"""
        try:
            while not STATUS_QUEUE.empty():
                msg = STATUS_QUEUE.get_nowait()
                self.status_label.config(text=f"当前状态：{msg}")
        except:
            pass
        self.root.after(100, self.consume_status_queue)

    def run(self):
        """启动悬浮窗"""
        logger.info("悬浮窗启动成功：最终最终版，解决返回按钮SVG点击问题")
        self.root.mainloop()


# ===================== 主函数 =====================
if __name__ == "__main__":
    logger.info("=" * 50)
    logger.info("自动学习程序启动（最终最终版：解决SVG返回按钮点击）")
    logger.info(f"Python版本：{sys.version}")
    logger.info("=" * 50)

    # 启动悬浮窗
    float_window = TaskFloatWindow()
    float_window.run()

    logger.info("程序已退出")
    sys.exit(0)