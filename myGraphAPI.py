import requests
from requests import Response
import json
import base64


HOST = "https://graph.microsoft.com/v1.0"

# Microsoft Graph API Permissions (Application permissions):
# - User.ReadBasic.All
# - Application.ReadAll
# - Tasks.ReadWrite.All
# - Mail.Send


# =============================== Authentication and authorization =============================== #


def organize_result(response: Response, expected_status_code: int) -> dict:
    """整理请求结果
    :param response: 请求的响应结果
    :param expected_status_code: 期望的状态码
    :return: 整理后的返回结果
    """
    result = {
        "is_success": False,
        "status_code": 0,
        "expected_status_code": expected_status_code,
        "response": {},
        "response_text": "",
    }
    result["status_code"] = response.status_code
    result["response_text"] = response.text
    if (
        "Content-Type" in response.headers
        and "application/json" in response.headers["Content-Type"]
    ):
        result["response"] = response.json()
    if response.status_code == result["expected_status_code"]:
        result["is_success"] = True
    return result


def get_credential(tenant: str, client_id: str, client_secret: str) -> dict:
    """获取凭证 (非登录认证)
    :param tenant: 租户名称, 通常为"xxx.onmicrosoft.com"
    :param client_id: 应用程序ID (Application ID, appId), 又名客户端ID (Client ID)
    :param client_secret: 应用程序密码
    :return: 整理后的返回结果
    """
    url = f"https://login.microsoftonline.com/{tenant}/oauth2/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "resource": "https://graph.microsoft.com",
    }
    response: Response = requests.post(url=url, data=data)
    return organize_result(response=response, expected_status_code=200)


# =============================== Authentication and authorization =============================== #


# ======================================== Users -> User ========================================= #
# Permission: User.ReadBasic.All


## Users -> User -> List
def list_users(token: str) -> dict:
    """列出所有用户
    :param token: 访问令牌
    :return: 整理后的返回结果
    """
    url = f"{HOST}/users"
    headers = {"Authorization": f"Bearer {token}"}
    return organize_result(
        response=requests.get(url=url, headers=headers), expected_status_code=200
    )


# ======================================== Users -> User ========================================= #

# ================================= Applications -> Application ================================== #
# Permission: Application.ReadAll


## Applications -> Application -> Owners -> List
def list_owners(token: str, client_id: str) -> dict:
    """列出指定应用的所有拥有者
    :param token: 访问令牌
    :param client_id: 客户端ID
    :return: 整理后的返回结果
    """
    url = f"{HOST}/applications(appId='{client_id}')/owners"
    headers = {"Authorization": f"Bearer {token}"}
    return organize_result(
        response=requests.get(url=url, headers=headers), expected_status_code=200
    )


# ================================= Applications -> Application ================================== #

# ========================================= To-do tasks ========================================== #
# Permission: Tasks.ReadWrite.All


## To-do tasks -> To-do task list -> List task lists
def list_task_lists(token: str, userId: str) -> dict:
    """列出所有任务列表
    :param token: 访问令牌
    :param userId: 用户ID
    :return: 整理后的返回结果
    """
    url = f"{HOST}/users/{userId}/todo/lists"
    headers = {"Authorization": f"Bearer {token}"}
    return organize_result(
        response=requests.get(url=url, headers=headers), expected_status_code=200
    )


## To-do tasks -> To-do task list -> Create task lists
def create_task_list(token: str, userId: str, data: dict) -> dict:
    """创建一个任务列表
    :param token: 访问令牌
    :param userId: 用户ID
    :param data: 任务列表数据
    :return: 整理后的返回结果
    """
    url = f"{HOST}/users/{userId}/todo/lists"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    return organize_result(
        response=requests.post(url=url, headers=headers, data=json.dumps(data)),
        expected_status_code=201,
    )


## To-do tasks -> To-do task list -> Get task lists
def get_task_list(token: str, userId: str, todoTaskListId: str) -> dict:
    """读取指定任务列表的信息
    :param token: 访问令牌
    :param userId: 用户ID
    :todoTaskListId: 任务列表ID
    :return: 整理后的返回结果
    """
    url = f"{HOST}/users/{userId}/todo/lists/{todoTaskListId}"
    headers = {"Authorization": f"Bearer {token}"}
    return organize_result(
        response=requests.get(url=url, headers=headers), expected_status_code=200
    )


## To-do tasks -> To-do task list -> Update task lists
def update_task_list(token: str, userId: str, todoTaskListId: str, data: dict) -> dict:
    """更新指定任务列表
    :param token: 访问令牌
    :param userId: 用户ID
    :param todoTaskListId: 任务列表ID
    :param data: 任务列表数据
    :return: 整理后的返回结果
    """
    url = f"{HOST}/users/{userId}/todo/lists/{todoTaskListId}"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    return organize_result(
        response=requests.patch(url=url, headers=headers, data=json.dumps(data)),
        expected_status_code=200,
    )


## To-do tasks -> To-do task list -> Delete task lists
def delete_task_list(token: str, userId: str, todoTaskListId: str) -> dict:
    """删除指定任务列表
    :param token: 访问令牌
    :param userId: 用户ID
    :todoTaskListId: 任务列表ID
    :return: 整理后的返回结果
    """
    url = f"{HOST}/users/{userId}/todo/lists/{todoTaskListId}"
    headers = {"Authorization": f"Bearer {token}"}
    return organize_result(
        response=requests.delete(url=url, headers=headers), expected_status_code=204
    )


## To-do tasks -> To-do task -> List tasks
def list_tasks(token: str, userId: str, todoTaskListId: str) -> dict:
    """列出指定任务列表下的所有任务
    :param token: 访问令牌
    :param userId: 用户ID
    :param todoTaskListId: 任务列表ID
    :return: 整理后的返回结果
    """
    url = f"{HOST}/users/{userId}/todo/lists/{todoTaskListId}/tasks"
    headers = {"Authorization": f"Bearer {token}"}
    return organize_result(
        response=requests.get(url=url, headers=headers), expected_status_code=200
    )


## To-do tasks -> To-do task -> Create task
def create_task(token: str, userId: str, todoTaskListId: str, data: dict) -> dict:
    """创建一个指定任务列表下的任务
    :param token: 访问令牌
    :param userId: 用户ID
    :param todoTaskListId: 任务列表ID
    :param data: 任务数据
    :return: 整理后的返回结果
    """
    url = f"{HOST}/users/{userId}/todo/lists/{todoTaskListId}/tasks"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    return organize_result(
        response=requests.post(url=url, headers=headers, data=json.dumps(data)),
        expected_status_code=201,
    )


## To-do tasks -> To-do task -> Get task
def get_task(token: str, userId: str, todoTaskListId: str, todoTaskId: str) -> dict:
    """读取指定任务列表下指定任务的信息
    :param token: 访问令牌
    :param userId: 用户ID
    :param todoTaskListId: 任务列表ID
    :param todoTaskId: 任务ID
    :return: 整理后的返回结果
    """
    url = f"{HOST}/users/{userId}/todo/lists/{todoTaskListId}/tasks/{todoTaskId}"
    headers = {"Authorization": f"Bearer {token}"}
    return organize_result(
        response=requests.get(url=url, headers=headers), expected_status_code=200
    )


## To-do tasks -> To-do task -> Update task
def update_task(
    token: str, userId: str, todoTaskListId: str, todoTaskId: str, data: dict
) -> dict:
    """更新指定任务列表下的指定任务
    :param token: 访问令牌
    :param userId: 用户ID
    :param todoTaskListId: 任务列表ID
    :param todoTaskId: 任务ID
    :param data: 任务数据
    :return: 整理后的返回结果
    """
    url = f"{HOST}/users/{userId}/todo/lists/{todoTaskListId}/tasks/{todoTaskId}"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    return organize_result(
        response=requests.patch(url=url, headers=headers, data=json.dumps(data)),
        expected_status_code=200,
    )


## To-do tasks -> To-do task -> Delete task
def delete_task(token: str, userId: str, todoTaskListId: str, todoTaskId: str) -> dict:
    """删除指定任务列表下的指定任务
    :param token: 访问令牌
    :param userId: 用户ID
    :param todoTaskListId: 任务列表ID
    :param todoTaskId: 任务ID
    :return: 整理后的返回结果
    """
    url = f"{HOST}/users/{userId}/todo/lists/{todoTaskListId}/tasks/{todoTaskId}"
    headers = {"Authorization": f"Bearer {token}"}
    return organize_result(
        response=requests.delete(url=url, headers=headers), expected_status_code=204
    )


## To-do tasks -> Checklist item -> List
def list_checklist_items(
    token: str, userId: str, todoTaskListId: str, todoTaskId: str
) -> dict:
    """列出指定任务列表下指定任务的所有检查项
    :param token: 访问令牌
    :param userId: 用户ID
    :param todoTaskListId: 任务列表ID
    :param todoTaskId: 任务ID
    :return: 整理后的返回结果
    """
    url = f"{HOST}/users/{userId}/todo/lists/{todoTaskListId}/tasks/{todoTaskId}/checklistItems"
    headers = {"Authorization": f"Bearer {token}"}
    return organize_result(
        response=requests.get(url=url, headers=headers), expected_status_code=200
    )


## To-do tasks -> Checklist item -> Create
def create_checklist_item(
    token: str, userId: str, todoTaskListId: str, todoTaskId: str, data: str
) -> dict:
    """创建一个指定任务列表下指定任务的检查项
    :param token: 访问令牌
    :param userId: 用户ID
    :param todoTaskListId: 任务列表ID
    :param todoTaskId: 任务ID
    :param data: 检查项数据
    :return: 整理后的返回结果
    """
    url = f"{HOST}/users/{userId}/todo/lists/{todoTaskListId}/tasks/{todoTaskId}/checklistItems"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    return organize_result(
        response=requests.post(url=url, headers=headers, data=json.dumps(data)),
        expected_status_code=201,
    )


## To-do tasks -> Checklist item -> Get
def get_checklist_item(
    token: str, userId: str, todoTaskListId: str, todoTaskId: str, checklistItemId: str
) -> dict:
    """读取一个指定任务列表下指定任务中指定检查项的信息
    :param token: 访问令牌
    :param userId: 用户ID
    :param todoTaskListId: 任务列表ID
    :param todoTaskId: 任务ID
    :param checklistItemId: 检查项ID
    :param data: 检查项数据
    :return: 整理后的返回结果
    """
    url = f"{HOST}/users/{userId}/todo/lists/{todoTaskListId}/tasks/{todoTaskId}/checklistItems/{checklistItemId}"
    headers = {"Authorization": f"Bearer {token}"}
    return organize_result(
        response=requests.get(url=url, headers=headers), expected_status_code=200
    )


## To-do tasks -> Checklist item -> Update
def update_checklist_item(
    token: str,
    userId: str,
    todoTaskListId: str,
    todoTaskId: str,
    checklistItemId: str,
    data: str,
) -> dict:
    """更新指定任务列表下指定任务的指定检查项
    :param token: 访问令牌
    :param userId: 用户ID
    :param todoTaskListId: 任务列表ID
    :param todoTaskId: 任务ID
    :param checklistItemId: 检查项ID
    :param data: 检查项数据
    :return: 整理后的返回结果
    """
    url = f"{HOST}/users/{userId}/todo/lists/{todoTaskListId}/tasks/{todoTaskId}/checklistItems/{checklistItemId}"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    return organize_result(
        response=requests.patch(url=url, headers=headers, data=json.dumps(data)),
        expected_status_code=200,
    )


## To-do tasks -> Checklist item -> Delete
def delete_checklist_item(
    token: str, userId: str, todoTaskListId: str, todoTaskId: str, checklistItemId: str
) -> dict:
    """删除指定任务列表下指定任务的指定检查项
    :param token: 访问令牌
    :param userId: 用户ID
    :param todoTaskListId: 任务列表ID
    :param todoTaskId: 任务ID
    :param checklistItemId: 检查项ID
    :return: 整理后的返回结果
    """
    url = f"{HOST}/users/{userId}/todo/lists/{todoTaskListId}/tasks/{todoTaskId}/checklistItems/{checklistItemId}"
    headers = {"Authorization": f"Bearer {token}"}
    return organize_result(
        response=requests.delete(url=url, headers=headers), expected_status_code=204
    )


# ========================================= To-do tasks ========================================== #

# ==================================== Users -> User -> Mail ===================================== #
# Permission: Mail.Send


## Users -> User -> Mail -> Send mail
def send_mail_json(
    token: str, userId: str, message: dict, saveToSentItems: bool = True
) -> dict:
    """发送邮件 (JSON/字典格式)
    :param token: 访问令牌
    :param userId: 发件者用户ID
    :param message: JSON/字典格式邮件数据
    :param saveToSentItems: 是否保存到发件人的SentItems文件夹 (默认为True)
    :return: 整理后的返回结果
    """
    url = f"{HOST}/users/{userId}/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    data = {
        "message": message,
        "saveToSentItems": saveToSentItems,
    }
    return organize_result(
        response=requests.post(url=url, headers=headers, data=json.dumps(data)),
        expected_status_code=202,
    )


## Users -> User -> Mail -> Send mail
def send_mail_mime(token: str, userId: str, mime: str) -> dict:
    """发送邮件 (MIME格式)
    :param token: 访问令牌
    :param userId: 发件者用户ID
    :param mime: MIME格式邮件数据
    :return: 整理后的返回结果
    """
    url = f"{HOST}/users/{userId}/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "text/plain"}
    data = base64.b64encode(mime.encode("utf-8")).decode("utf-8")
    return organize_result(
        response=requests.post(url=url, headers=headers, data=data),
        expected_status_code=202,
    )


# ==================================== Users -> User -> Mail ===================================== #
