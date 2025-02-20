<div align="center">
    <img src="https://socialify.git.ci/imoki/wpsPyInf/image?description=1&font=Rokkitt&forks=1&issues=1&language=1&owner=1&pattern=Circuit%20Board&pulls=1&stargazers=1&theme=Dark">
<h1>默代理</h1>
基于「金山文档」的python任务代理

<div id="shield">

[![][github-stars-shield]][github-stars-link]
[![][github-forks-shield]][github-forks-link]
[![][github-issues-shield]][github-issues-link]
[![][github-contributors-shield]][github-contributors-link]

<!-- SHIELD GROUP -->
</div>
</div>

## 👑 背景
在金山文档中  
1. 由于一个文档内定时任务限制不超过十个，导致很多人要设置超过10个任务时就必须多建几个文档，并且一个个配置比较繁琐。  
2. 其次是目前官方还未支持python定时任务，导致想设置定时需要用airscript开发来设置定时  
3. 最后是python代码只能在编辑器内写  

## 🎊 简介
本项目另辟蹊径解决了以上问题。  
1. 一个文档内可以设置无限个定时任务
2. 使用python代码的定时任务
3. python代码直接在表格中写
4. 可接受python接受传参数
5. 可接入消息推送
6. 可自动更新脚本
7. 支持netcut、github代理、github脚本下载

## ✨ 特性
    - 📀 支持金山文档运行
    - 💿 支持普通表格和智能表格
    - ♾️ 无限制python脚本数量
    - 💽 定时任务随意设置
    - 💿 支持脚本自动更新
    - 💿 支持github脚本自动下载
    
## 🍨 教程说明
💬 公众号“默库”

## 🛰️ 文字步骤
注意！请将文档名和脚本名都起名为“默库”，脚本才能正常运行。  
1. 第一步，首次运行“默库”脚本（仓库中的“moku.js”）会生成wps表，请先填写好wps表的内容，只填wps_sid即可。
2. 填写CONFIG表的内容。默认任务1用于消息推送测试，测试脚本是否正常，填写推送的key即可，如:bark=xxxx&pushplus=xxxx。(这一步可以跳过)
3. 再运行一次“默库”脚本，此时你将收到推送通知，说明你操作正确，可正常使用了。(这一步可以跳过)
4. 请在CONFIG表填写你自己写的python脚本和定时时间，然后运行一次“默库”脚本，即可按照配置好的来执行脚本，就不需要再管了。

## 🛰️ CONFIG表内容
![](https://s3.bmp.ovh/imgs/2024/07/26/d4f569687ade2c29.png)

## 🛰️ 远程脚本下载脚本模式
1. netcut   --  从netcut.cn中下载，如：https://netcut.cn/p/9aa97e54eb186c06
2. githubproxy  --  从github代理下载，如：https://raw.kkgithub.com/imoki/wpsPyInf/main/testPush.py
3. github   --  从github直接下载，如：https://github.com/imoki/wpsPyInf/blob/main/testPush.py

## 🛰️ 动态更新远程的python脚本如何接入此项目
1. 请将python脚本写入netcut中
```
https://netcut.cn/
```  
2. 在python脚本中用如下方式标识唯一id，可写在最开头。  
```python
uniqueId = "xxxx"
```  
项目会自动获取此值写入表格中，表格中有“唯一id”列，根据此值来判断是否是脚本所处表格的行。   
2. 在对应的“脚本地址”中写入分享地址  

## 🛰️ 如何接入消息推送
请参考这个项目: https://github.com/imoki/wpsPush  

## 🤝 欢迎参与贡献
欢迎各种形式的贡献

[![][pr-welcome-shield]][pr-welcome-link]

<!-- ### 💗 感谢我们的贡献者
[![][github-contrib-shield]][github-contrib-link] -->


## ✨ Star 数

[![][starchart-shield]][starchart-link]

## 📝 更新日志 
- 2024-07-26
    * 增加从github直链下载脚本的功能
    * 增加“是否禁止更新”选项
- 2024-07-24
    * 修复时间写入和取出表格时被金山自动进行格式化的问题，防止重复更新脚本
    * 增加从github代理下载脚本的功能
- 2024-07-21
    * 支持远程脚本动态更新
- 2024-07-20
    * 推出默代理

## 📌 特别声明

- 本仓库发布的脚本仅用于测试和学习研究，禁止用于商业用途，不能保证其合法性，准确性，完整性和有效性，请根据情况自行判断。

- 本人对任何脚本问题概不负责，包括但不限于由任何脚本错误导致的任何损失或损害。

- 间接使用脚本的任何用户，包括但不限于建立VPS或在某些行为违反国家/地区法律或相关法规的情况下进行传播, 本人对于由此引起的任何隐私泄漏或其他后果概不负责。

- 请勿将本仓库的任何内容用于商业或非法目的，否则后果自负。

- 如果任何单位或个人认为该项目的脚本可能涉嫌侵犯其权利，则应及时通知并提供身份证明，所有权证明，我们将在收到认证文件后删除相关脚本。

- 任何以任何方式查看此项目的人或直接或间接使用该项目的任何脚本的使用者都应仔细阅读此声明。本人保留随时更改或补充此免责声明的权利。一旦使用并复制了任何相关脚本或Script项目的规则，则视为您已接受此免责声明。

**您必须在下载后的24小时内从计算机或手机中完全删除以上内容**

> ***您使用或者复制了本仓库且本人制作的任何脚本，则视为 `已接受` 此声明，请仔细阅读***

<!-- LINK GROUP -->
[github-codespace-link]: https://codespaces.new/imoki/wpsPyInf
[github-codespace-shield]: https://github.com/imoki/wpsPyInf/blob/main/images/codespaces.png?raw=true
[github-contributors-link]: https://github.com/imoki/wpsPyInf/graphs/contributors
[github-contributors-shield]: https://img.shields.io/github/contributors/imoki/wpsPyInf?color=c4f042&labelColor=black&style=flat-square
[github-forks-link]: https://github.com/imoki/wpsPyInf/network/members
[github-forks-shield]: https://img.shields.io/github/forks/imoki/wpsPyInf?color=8ae8ff&labelColor=black&style=flat-square
[github-issues-link]: https://github.com/imoki/wpsPyInf/issues
[github-issues-shield]: https://img.shields.io/github/issues/imoki/wpsPyInf?color=ff80eb&labelColor=black&style=flat-square
[github-stars-link]: https://github.com/imoki/wpsPyInf/stargazers
[github-stars-shield]: https://img.shields.io/github/stars/imoki/wpsPyInf?color=ffcb47&labelColor=black&style=flat-square
[github-releases-link]: https://github.com/imoki/wpsPyInf/releases
[github-releases-shield]: https://img.shields.io/github/v/release/imoki/wpsPyInf?labelColor=black&style=flat-square
[github-release-date-link]: https://github.com/imoki/wpsPyInf/releases
[github-release-date-shield]: https://img.shields.io/github/release-date/imoki/wpsPyInf?labelColor=black&style=flat-square
[pr-welcome-link]: https://github.com/imoki/wpsPyInf/pulls
[pr-welcome-shield]: https://img.shields.io/badge/🤯_pr_welcome-%E2%86%92-ffcb47?labelColor=black&style=for-the-badge
[github-contrib-link]: https://github.com/imoki/wpsPyInf/graphs/contributors
[github-contrib-shield]: https://contrib.rocks/image?repo=imoki%2Fsign_script
[docker-pull-shield]: https://img.shields.io/docker/pulls/imoki/wpsPyInf?labelColor=black&style=flat-square
[docker-pull-link]: https://hub.docker.com/repository/docker/imoki/wpsPyInf
[docker-size-shield]: https://img.shields.io/docker/image-size/imoki/wpsPyInf?labelColor=black&style=flat-square
[docker-size-link]: https://hub.docker.com/repository/docker/imoki/wpsPyInf
[docker-stars-shield]: https://img.shields.io/docker/stars/imoki/wpsPyInf?labelColor=black&style=flat-square
[docker-stars-link]: https://hub.docker.com/repository/docker/imoki/wpsPyInf
[starchart-shield]: https://api.star-history.com/svg?repos=imoki/wpsPyInf&type=Date
[starchart-link]: https://api.star-history.com/svg?repos=imoki/wpsPyInf&type=Date

