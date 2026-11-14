# -*- coding: utf-8 -*-
"""版本号唯一来源。

发版流程：改这里的 __version__ -> 提交 -> 打对应的 vX.Y.Z 标签推送。
Release 工作流会先校验标签与本文件一致，再跑测试、构建发布；
CLI 的 --version、GUI 窗口标题、exe 版本资源都从这里取值。
"""

__version__ = '1.3.0'
