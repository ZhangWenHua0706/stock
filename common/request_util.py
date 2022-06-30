import requests


class RequestUtil:
    session = requests.session()

    def send_request(self, method, url, data, post_type, **kwargs):
        proxies = {
            "https":"https://127.0.0.1:53385",
            "http":"http://127.0.0.1:53385"
        }
        method = str(method).lower()
        if method == 'get':
            print("ffff")
            #rep = RequestUtil.session.request(method, url=url, params=data, **kwargs)
            rep = RequestUtil.session.request(method, url=url, params=data, **kwargs,proxies=proxies,verify=False)

        elif method == 'post':
            if post_type == "form":
                rep = RequestUtil.session.request(method, url=url, data=data, **kwargs,proxies=proxies,verify=False)
                #rep = RequestUtil.session.request(method, url=url, data=data, **kwargs)
            elif post_type == "json":
                #data = json.dumps(data, ensure_ascii=False)
                rep = RequestUtil.session.request(method, url=url, json=data, **kwargs,proxies=proxies,verify=False)
                #rep = RequestUtil.session.request(method, url=url, json=data, **kwargs)
            elif post_type == "file":
                #data = json.dumps(data, ensure_ascii=False)
                rep = RequestUtil.session.request(method, url=url, files=data, **kwargs,proxies=proxies,verify=False)
                #rep = RequestUtil.session.request(method, url=url, json=data, **kwargs)
            else:
                return "post type error"
        else:
            return "Method error"
        rep.encoding = 'UTF-8'

        return rep.text
