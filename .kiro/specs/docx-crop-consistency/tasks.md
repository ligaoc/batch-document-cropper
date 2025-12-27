# Implementation Plan: DOCX Crop Consistency

## Overview

修改 `DOCXCropper` 类，将裁剪逻辑从"设置边距"改为"累加边距"，确保裁剪结果与预览一致。

## Tasks

- [x] 1. 修改 DOCXCropper 类实现边距累加逻辑
  - [x] 1.1 添加 `_get_section_margin()` 方法读取节的原始边距
    - 处理边距为 None 的情况，返回默认值 25.4mm
    - _Requirements: 2.1, 2.2_
  - [x] 1.2 添加 `_validate_new_margins()` 方法验证新边距
    - 检查左右边距之和不超过页面宽度
    - 检查上下边距之和不超过页面高度
    - 返回错误信息指明哪个边距无效
    - _Requirements: 4.1, 4.2, 4.3_
  - [x] 1.3 修改 `crop()` 方法实现边距累加
    - 对每个节：读取原始边距 → 计算新边距 → 验证 → 应用
    - 新边距 = 原始边距 + 裁剪量
    - _Requirements: 1.1, 1.2, 1.3, 1.4_

- [ ] 2. 添加单元测试
  - [ ]* 2.1 测试边距累加计算正确性
    - 测试原始边距 25mm + 裁剪量 28mm = 53mm
    - 测试裁剪量为 0 时边距不变
    - _Requirements: 1.1, 1.2, 1.3_
  - [ ]* 2.2 测试边距验证逻辑
    - 测试超出页面尺寸时返回错误
    - 测试错误信息包含边距信息
    - _Requirements: 4.1, 4.2_

- [ ] 3. Checkpoint - 确保所有测试通过
  - 运行测试验证实现正确性
  - 如有问题请询问用户

## Notes

- 任务标记 `*` 为可选测试任务
- 主要改动集中在 `src/core/docx_cropper.py`
- 无需修改其他组件，因为 `MarginSettings` 的数据结构不变
