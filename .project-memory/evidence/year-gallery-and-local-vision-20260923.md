# 报告图片年份对比与本地视觉复核

日期：2026-09-23

后续修复批次已完成，最终结论见 [corrected-real-batch-20260923.md](corrected-real-batch-20260923.md)。下文保留首轮问题发现与当时的验证状态。

## 报告页

统一处理 HTML 报告现从原稿和输出 DOCX 中按 `word/media` 路径配对图片，展示发生像素变化的完整图与局部放大图，并明确标为“待核验”。原稿位置使用报告的 `material_root`，输出位置使用 `output`；预览阶段不生成对比。资源放在 HTML 同目录的 `_assets` 文件夹，以相对路径引用，可离线打开。

旧全量批次 `.omc/full-batch-work-v2/combined_document_result.json` 重新生成 HTML 后，出现 10 个对比卡片：医院 2、建筑 1、毫米波 7，与日期桥接 `changed_images` 合计一致。所有 60 个本地图片引用存在。独立逐图审核发现其中仅医院 `image19.png` 的日历标题 `2022→2025` 可确认为真实年份；医院 `image20.png` 是人员编号误改，建筑 `image11.png` 是图表刻度误改，毫米波 7 张是时间列误改。**旧批次的 9 张误改图不是正确成果，旧输出不能直接交付。** Word 后续缩放造成的 14 张非年份图片差异已按尺寸变化排除。

## 本地视觉接口

默认路由为 `http://127.0.0.1:3458/v1/chat/completions`，模型固定为 `sensenova-6.8-flash-lite`，密钥使用本地占位值 `proxy`，也可由环境变量配置 Base URL 和 Key。每项目最多对 2 张 OCR `unknown`/`irrelevant` 候选等距抽查，单次请求限时 20 秒；失败记为 `unknown`。模型返回的视觉判断单独写入 `visual_review`，不进入 OCR `blocked`、批量熔断或日期写入规则。

真实 10 份说明书预览：`.omc/verification/vision_real10_preview_v2.json`。20 次调用得到 `relevant=12`、`irrelevant=1`、`unknown=7`（其中接口失败 4），共 25,861 tokens。太阳能 LED 路灯储能说明书有一张图片被判不相关；原有 OCR 拦截仍为 0，批量熔断仍为否。视觉抽样未覆盖或未确认的图片不能据此认定相关。

## 验证

首轮 `python\python.exe -m unittest discover -s tests -q`：123 项通过，但未覆盖上述 9 张误改；日期桥接正在增加真实图回归，修复结果以追加验证为准。最终 RAR 未打包。
