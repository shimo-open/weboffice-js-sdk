# API Contract Catalog

这里是 JS-API 的机器可读契约目录，也是新增 API 的唯一登记入口。

新增 API 时必须：

1. 在对应 suite JSON 中登记；
2. 运行 `npm run contract-check`；
3. 同步 iframe receiver、typed facade、类型断言、测试和文档；Showcase 属于业务层，按业务项目需要同步；
4. 通过 `npm test`、`npm run lint` 和 TypeScript 检查。

胶水层只负责接口调用、通信、数据适配和错误处理，不负责权限判断、安全策略、多个接口的业务编排或修改结果验收。需要验证业务闭环时，在套件能力层或业务项目中运行：

```bash
npm run showcase-contract-check
```

详细流程见 [`doc/feature/jsapi-extension-playbook.md`](../doc/feature/jsapi-extension-playbook.md)。
