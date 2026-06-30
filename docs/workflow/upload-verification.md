# 上传核验记录

## 核验时间

2026-06-30

## 远端仓库

https://github.com/JT-6815/smart-product-development-magnetic-cargo-drone-validation

## 已上传提交

- 分支：`main`
- 提交：`43da0708bb43bec484844e201633a1595f95eeb8`
- 提交说明：`Archive smart product drone project materials`

## 归档来源

- `D:\大作业文件\大三下期\智能产品开发`
- `D:\大作业文件\大三下期\作品集设计\作品集作业\1230802221_张誉璁_智能产品开发设计\过程视频剪辑.mp4`
- `D:\大作业文件\大三下期\作品集设计\作品集作业\1230802221_张誉璁_智能产品开发设计\无人机结课报告.pptx`
- `D:\无人机`

## 仓库分类结果

| 仓库路径 | 文件数 | 内容说明 |
| --- | ---: | --- |
| `docs/` | 14 | 项目说明、工作流、上传索引、PPT 生成脚本与 QA 记录 |
| `images/` | 27 | 课程过程图片、起飞素材、组装与调试照片 |
| `materials/` | 158 | F450、APM、Pixhawk、Mission Planner、固件、驱动、教程视频等参考资料 |
| `reports/` | 9 | 课程 PPT、结课报告 PPT、作品集 PPT 与成员记录 |
| `videos/` | 1 | 作品集过程视频剪辑 |
| `cad/` | 2 | 结构与磁吸模块相关资料 |
| `tests/` | 3 | 测试数据模板与验证记录 |

仓库当前共纳入 Git 管理文件 `217` 个。

## Git LFS 状态

本次上传使用 Git LFS 管理大体积或二进制资料，包括 PPTX、MP4、MOV、RAR、ZIP、EXE、MSI、DOC、PDF、APK、BIN、PX4、JPG、PNG 等类型。

- LFS 跟踪文件数：`177`
- 已上传唯一 LFS 对象：`171`
- LFS 对象总体量：约 `7.2GB`

## 核验命令

```powershell
git status --short --branch
git ls-remote --heads origin main
git lfs ls-files | Measure-Object
git -c core.quotepath=false ls-files | Measure-Object
```

## 核验结论

本地 `main` 分支已与远端 `origin/main` 同步，远端 `main` 指向提交 `43da0708bb43bec484844e201633a1595f95eeb8`。所有分类文件已纳入版本控制，大文件已通过 Git LFS 上传完成。
