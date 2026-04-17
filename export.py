import asyncio
import os
from playwright.async_api import async_playwright

# --- 配置 ---
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
USER_DATA_DIR = os.path.join(os.getcwd(), "feishu_session")
BITABLE_URL = "https://mcn9cqf1t1jg.feishu.cn/base/Ntkeb5ZJhaeghvsa91DcO2Mdnog?table=tblUeHywI5GDE0PB&view=vewaUJNj7j"


async def export_process():
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

        print(f"[*] 跳转飞书...")
        await page.goto(BITABLE_URL, wait_until="commit")

        # 1. 等待 Head 栏并执行规约等待 10s
        await page.wait_for_selector(".note-title--header-flexible", timeout=60000)
        print("[*] Head 栏已加载，强制等待 10s 确保环境稳定...")
        await asyncio.sleep(10)

        try:
            # 2. 点击更多
            print("[*] 步骤 1: 点击【更多】")
            await page.locator('[data-selector="more-menu"]').first.click()
            await asyncio.sleep(2)

            # 3. 核心大招：坐标偏移点击（向右 80px）
            print("[*] 步骤 2: 定位‘导出’项并执行 80px 偏移点击...")
            # 找到“导出”这个文字所在的菜单行
            export_item = page.locator('div[role="menuitem"]:has-text("导出")').first
            box = await export_item.bounding_box()

            if box:
                # 计算目标位置：导出项中心点 + 向右 80px
                target_x = box['x'] + box['width'] / 2 + 180
                target_y = box['y'] + box['height'] / 2

                # 模拟鼠标轨迹：先移动到“导出”，再平移点击
                await page.mouse.move(box['x'] + box['width'] / 2, box['y'] + box['height'] / 2)
                await asyncio.sleep(0.5)
                await page.mouse.click(target_x, target_y)
                print(f"[+] 偏移点击成功 (坐标偏移: +80px)")
            else:
                print("[X] 无法获取‘导出’项坐标，请检查菜单是否弹出")
                return

            # 4. 等待“下载设置”弹窗出现
            print("[*] 步骤 3: 等待下载设置弹窗...")
            await page.wait_for_selector(".export-setting-modal", timeout=15000)

            # 5. 处理虚拟列表下拉框
            print("[*] 步骤 4: 展开‘下载的数据范围’下拉框")
            await page.locator(".export-setting-select").click()
            await asyncio.sleep(2)  # 虚拟列表渲染需要一点时间

            # 6. 精准点击‘限额saas’
            print("[*] 步骤 5: 从虚拟列表中定位‘限额saas’...")
            # 根据你提供的 HTML 结构，使用精准选择器
            target_option = page.locator('.ud__select__list__item').filter(has_text="限额saas").first

            await target_option.scroll_into_view_if_needed()
            await target_option.click(force=True)
            print("[+] 已成功选中‘限额saas’")

            # 7. 最终执行下载
            print("[*] 步骤 6: 确认下载...")
            async with page.expect_download() as download_info:
                modal = page.locator(".export-setting-modal")
                # 寻找弹窗内蓝色的“下载”按钮
                await modal.locator('button:has-text("下载")').click()

            # 8. 保存文件
            download = await download_info.value
            save_path = os.path.join(os.getcwd(), download.suggested_filename)
            await download.save_as(save_path)

            print(f"\n{'=' * 30}")
            print(f"【导出成功】\n文件名: {download.suggested_filename}\n路径: {save_path}")
            print(f"{'=' * 30}")

        except Exception as e:
            print(f"\n[X] 脚本执行中断: {e}")
            await asyncio.sleep(20)  # 出错后留出观察时间
        finally:
            await context.close()


if __name__ == "__main__":
    asyncio.run(export_process())