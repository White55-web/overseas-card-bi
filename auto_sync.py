import os
import subprocess
import time
from datetime import datetime

# =================【需要监听的目录与文件配置】=================
WATCH_FOLDERS = ["Forest_Data", "Tim_Data"]
WATCH_DICTS = ["Forest字典.xlsx", "Tim字典.xlsx"]


def get_file_mtimes():
    """获取所有流水文件和字典文件的最后修改时间字典"""
    mtimes = {}

    # 1. 扫描流水文件夹
    for folder in WATCH_FOLDERS:
        if os.path.exists(folder):
            for fname in os.listdir(folder):
                # 过滤 Excel 临时锁文件 (~$)
                if fname.endswith(".xlsx") and not fname.startswith("~$"):
                    fpath = os.path.join(folder, fname)
                    try:
                        mtimes[fpath] = os.path.getmtime(fpath)
                    except Exception:
                        pass

    # 2. 扫描两个字典文件
    for dict_name in WATCH_DICTS:
        if os.path.exists(dict_name):
            try:
                mtimes[dict_name] = os.path.getmtime(dict_name)
            except Exception:
                pass

    return mtimes


def push_to_github(trigger_file):
    """触发 Git 自动提交与推送"""
    fname = os.path.basename(trigger_file)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if "字典" in fname:
        commit_msg = f"🤖 Auto-sync: 字典更新 [{fname}] - {now_str}"
    else:
        commit_msg = f"🤖 Auto-sync: {fname}"

    print(f"[{now_str}] ⚡ 检测到文件变动: {fname}，正在自动上传 Git...")

    try:
        # 1. 添加所有变动（包含流水文件夹与根目录字典）
        subprocess.run(["git", "add", "."], check=True)

        # 2. 提交 commit
        subprocess.run(["git", "commit", "-m", commit_msg], check=True)

        # 3. 推送到远程分支
        subprocess.run(["git", "push"], check=True)
        print(f"[{now_str}]  成功推送至 GitHub: {fname}")
    except subprocess.CalledProcessError as e:
        print(f"[{now_str}] ⚠️ 推送失败或无内容更新: {e}")
    except Exception as e:
        print(f"[{now_str}] ❌ 发生异常: {e}")


def main():
    print("🚀 自动化中台同步引擎已启动...")
    print(f"📁 监听流水文件夹: {WATCH_FOLDERS}")
    print(f" 监听字典文件: {WATCH_DICTS}")
    print("--------------------------------------------------")

    # 记录初始文件状态
    last_mtimes = get_file_mtimes()

    while True:
        try:
            time.sleep(3)  # 每 3 秒巡检一次
            current_mtimes = get_file_mtimes()

            # 检查是否有新增或修改的文件
            changed_file = None
            for fpath, mtime in current_mtimes.items():
                if fpath not in last_mtimes or mtime > last_mtimes[fpath]:
                    changed_file = fpath
                    break

            if changed_file:
                # 稍微等待 1 秒，确保 Excel 文件保存完毕写入完整
                time.sleep(1)
                push_to_github(changed_file)
                # 更新状态
                last_mtimes = get_file_mtimes()

        except KeyboardInterrupt:
            print("\n⏹️ 自动同步已停止。")
            break
        except Exception as e:
            time.sleep(3)


if __name__ == "__main__":
    main()