import asyncio
import os
from playwright.async_api import async_playwright

# --- 配置 ---
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
# 获取当前脚本所在目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USER_DATA_DIR = os.path.join(BASE_DIR, "tencent_session")
DOCS_URL = "https://docs.qq.com/sheet/DTFdkY3NqamJJVEJl?tab=1sweh3"


async def export_tencent_process(save_dir=None, target_filename=None):
    if not save_dir:
        save_dir = os.path.join(BASE_DIR, "resource")
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)

    if not os.path.exists(USER_DATA_DIR):
        os.makedirs(USER_DATA_DIR)

    async with async_playwright() as p:
        print(f"[*] 正在启动浏览器...")
        context = await p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            executable_path=CHROME_PATH,
            headless=False,
            args=['--start-maximized']
        )

        page = await context.new_page()
        page.set_default_timeout(60000)

        print(f"[*] 跳转腾讯文档...")
        await page.goto(DOCS_URL)

        # 1. 检查是否需要登录
        print("[*] 正在多维度检查是否需要登录...")
        
        async def try_click_login(frame):
            try:
                btn = frame.locator("#header-login-btn")
                if await btn.is_visible(timeout=2000):
                    await btn.click()
                    return True
                btn_text = frame.get_by_text("登录腾讯文档")
                if await btn_text.is_visible(timeout=1000):
                    await btn_text.first.click()
                    return True
            except:
                pass
            return False

        login_clicked = await try_click_login(page)
        if not login_clicked:
            for frame in page.frames:
                if await try_click_login(frame):
                    login_clicked = True
                    break
        
        if login_clicked:
            print("[!] 请在浏览器窗口中完成登录操作...")
        
        # 2. 等待文档加载
        print("[*] 正在等待文档加载...")
        try:
            await page.wait_for_selector("#main-menu-file", state="visible", timeout=300000) 
            print("[+] 文档加载成功")
        except Exception as e:
            print(f"[X] 等待超时: {e}")
            return None

        try:
            # 3. 点击菜单
            await page.locator("#main-menu-file").first.click()
            await asyncio.sleep(2)

            # 4. 触发下载
            async with page.expect_download() as download_info:
                export_btn = page.locator('.menu_workbench-menu-horizontal-item__1fd8H').filter(has_text="下载")
                await export_btn.first.click(force=True)

            # 5. 保存文件
            download = await download_info.value
            fname = target_filename if target_filename else download.suggested_filename
            save_path = os.path.join(save_dir, fname)
            await download.save_as(save_path)

            print(f"【腾讯文档导出成功】: {save_path}")
            return save_path

        except Exception as e:
            print(f"\n[X] 脚本中断: {e}")
            return None
        finally:
            await context.close()


if __name__ == "__main__":
    asyncio.run(export_tencent_process())
