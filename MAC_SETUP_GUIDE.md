# Mac Excel VBA 兼容性修复指南

## 🚨 问题根源
原有的 `JsonConverter.bas` 和 `Module_Config.bas` 使用了 `CreateObject("Scripting.Dictionary")`，这是Windows特有的ActiveX组件，在Mac Excel中无法使用。

## ✅ 解决方案
已创建完全Mac兼容的替代模块：

### 1. 新的Mac兼容模块
- **`Module_Config_Mac.bas`** - 替代 `Module_Config.bas`
- **`JsonConverter_Mac.bas`** - 替代 `JsonConverter.bas`  
- **`Module_ButtonHelper.bas`** - 按钮创建辅助工具

### 2. 导入步骤
1. 在VBA编辑器中导入这3个新模块
2. **不要删除**原有模块（保持兼容性）
3. 使用新的Mac兼容函数

### 3. 使用方法

#### 创建测试按钮
```vba
' 方法1：通过按钮辅助模块
Module_ButtonHelper.CreateMacConfigTestButton

' 方法2：直接调用
Module_ButtonHelper.QuickConfigToken
```

#### 配置API Token
```vba
' 使用Mac兼容配置模块
Module_Config_Mac.ShowConfigDialog
```

#### 测试系统
```vba
' 测试配置系统
Module_Config_Mac.TestConfigSystem

' 测试JSON转换器
JsonConverter_Mac.TestJSONConverter
```

### 4. 主要功能对照表

| 原有函数 | Mac兼容替代 | 功能 |
|---------|-------------|------|
| `Module_Config.ShowConfigDialog()` | `Module_Config_Mac.ShowConfigDialog()` | 配置对话框 |
| `Module_Config.SetConfig()` | `Module_Config_Mac.SetConfig()` | 设置配置 |
| `Module_Config.GetConfig()` | `Module_Config_Mac.GetConfig()` | 获取配置 |
| `JsonConverter.ParseJSON()` | `JsonConverter_Mac.ParseJSON()` | JSON解析 |
| `JsonConverter.ConvertToJSON()` | `JsonConverter_Mac.ConvertToJSON()` | JSON生成 |

### 5. 立即测试
运行以下命令测试所有功能：
```vba
Sub TestAllMacFeatures()
    ' 创建测试按钮
    Module_ButtonHelper.CreateMacConfigTestButton
    
    ' 测试配置系统
    Module_Config_Mac.TestConfigSystem
    
    ' 测试JSON转换
    JsonConverter_Mac.TestJSONConverter
    
    ' 检查模块状态
    Module_ButtonHelper.CheckModuleStatus
End Sub
```

## 🔧 配置文件格式
Mac兼容版本使用简单的key=value格式：
```
# API配置文件
api.token="你的API_TOKEN"
```

## 📝 注意事项
1. 新模块完全不依赖ActiveX组件
2. 配置文件自动保存到项目目录
3. 支持多种路径自动检测
4. 包含完整的错误处理和调试信息

现在可以在Mac Excel中正常使用所有配置功能！
