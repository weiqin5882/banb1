# 快手小店订单比对整理工具

一个纯前端网页工具，用于对比：
- 官方导出的订单表（WPS 导出）
- 客服手工统计订单表（WPS 导出）

## 功能

- 自动识别不同列名（如订单号/订单编号/快手订单编号等）
- 仅对比状态包含 `交易成功` 或 `已发货` 的订单（若某一张表没有状态列，会自动放宽为不过滤该表）
- 订单号比对，输出类序号、产品名称、销售额、成本、利润
- 自动标记漏单（官方缺客服 / 客服缺官方）
- 自动统计销售额合计、成本合计、总利润、亏损单数
- 亏损订单行红色高亮
- 支持导出 `订单比对结果.xlsx`

## 使用

1. 直接双击打开 `index.html`，或在本地起一个静态服务。
2. 选择官方订单表和客服统计表。
3. 可填写默认成本（如果导入表格没有成本列会使用它）。
4. 点击“开始比对”。
5. 点击“导出结果表格”。

## 服务器部署（推荐 Nginx）

> 这是纯前端项目，不依赖后端服务。只要能提供静态文件访问即可部署。

### 1）上传文件到服务器

把以下文件上传到服务器目录（例如 `/var/www/order-compare`）：

- `index.html`
- `app.js`
- `styles.css`

### 2）安装 Nginx

Ubuntu / Debian：

```bash
sudo apt update
sudo apt install -y nginx
```

CentOS / Rocky：

```bash
sudo yum install -y nginx
```

### 3）配置站点

新建配置文件（示例域名：`order.yourdomain.com`）：

```nginx
server {
    listen 80;
    server_name order.yourdomain.com;

    root /var/www/order-compare;
    index index.html;

    location / {
        try_files $uri $uri/ /index.html;
    }
}
```

启用并重载：

```bash
sudo nginx -t
sudo systemctl enable nginx
sudo systemctl restart nginx
```

### 4）开启防火墙端口

```bash
sudo ufw allow 80
sudo ufw allow 443
```

### 5）配置 HTTPS（可选但推荐）

```bash
sudo apt install -y certbot python3-certbot-nginx
sudo certbot --nginx -d order.yourdomain.com
```

证书自动续期检查：

```bash
sudo certbot renew --dry-run
```

## Docker 部署（可选）

在项目目录新增 `Dockerfile`（如需我也可以帮你补上）：

```dockerfile
FROM nginx:alpine
COPY index.html /usr/share/nginx/html/index.html
COPY app.js /usr/share/nginx/html/app.js
COPY styles.css /usr/share/nginx/html/styles.css
EXPOSE 80
```

构建并运行：

```bash
docker build -t order-compare:latest .
docker run -d --name order-compare -p 8080:80 order-compare:latest
```

然后访问：`http://服务器IP:8080`

## 支持的列名（可自动识别）

- 订单号：`快手订单编号`、`订单号`、`订单编号`、`订单ID`...
- 状态：`订单状态`、`状态`、`交易状态`...
- 产品：`订单商品名称`、`商品名称`、`产品名称`...
- 销售额：`商家实收`、`支付金额`、`订单金额`、`销售额`...
- 成本：`成本`、`采购价`、`采购成本`...

如果你的列名不在上述范围，建议在 `app.js` 中给 `headerAliases` 增加同义词。
