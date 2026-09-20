# API Contract Catalog

这里是 JS-API 的机器可读契约目录，也是新增 API 的唯一登记入口。

新增 API 时必须：

1. 在对应 suite JSON 中登记；
2. 如果 API 需要被 `sdk.canIUse()` 查询，在 `public-methods.json` 登记；
3. 运行 `npm run contract-check`；
4. 同步 iframe receiver、typed facade、类型断言、测试和文档；Showcase 属于业务层，按业务项目需要同步；
5. 通过 `npm test`、`npm run lint` 和 TypeScript 检查。

`public-methods.json` 中的 `publicPath` 是 SDK 导出的
`OfficeSDKMethods` 常量值，`productPath` 只供 iframe 映射使用。调用方不应手写
scope 字符串：

```ts
await sdk.canIUse(OfficeSDKMethods.ActiveOutline.Editor.Document.GetContent)
```

`sdk.canIUse()` 只判断当前套件、iframe runtime 和宿主 delegation 是否实现
该方法，不判断权限、选区或其他业务前置条件。未知方法返回 `false`；SDK 未连接
或通信失败仍然 reject。

胶水层只负责接口调用、通信、数据适配和错误处理，不负责权限判断、安全策略、多个接口的业务编排或修改结果验收。需要验证业务闭环时，在套件能力层或业务项目中运行：

```bash
npm run showcase-contract-check
```

详细流程见 [`doc/feature/jsapi-extension-playbook.md`](../doc/feature/jsapi-extension-playbook.md)。
