#include <stdio.h>
#include <string.h>
#include <iostream>
#include <tchar.h>
#include <windows.h>
#include <atlbase.h>
#include <UIAutomation.h>
#include <vector>
#include <locale.h>
using namespace std;

enum class WeChatMsgType {
	Unknown,       // 未知
	Notify,        // 系统通知
	Time,          // 时间
	Text,          // 文本
	Image,         // 图片
	Audio,         // 语音
	Video,         // 视频
	File,          // 文件
	Link,          // 链接
	RedPacket,     // 红包
	Transfer,      // 转账
	ChatRecord,    // 聊天记录
	Location,      // 地理位置
	VoiceCall,     // 音频通话
	VideoCall,     // 视频通话
};

std::wstring GetMsgTypeString(WeChatMsgType type) {
	switch (type)
	{
	case WeChatMsgType::Unknown:
		return L"Unknown";
	case WeChatMsgType::Notify:
		return L"通知";
	case WeChatMsgType::Time:
		return L"时间";
	case WeChatMsgType::Text:
		return L"文本";
	case WeChatMsgType::Image:
		return L"图片";
	case WeChatMsgType::Audio:
		return L"音频";
	case WeChatMsgType::Video:
		return L"视频";
	case WeChatMsgType::File:
		return L"文件";
	case WeChatMsgType::Link:
		return L"链接";
	case WeChatMsgType::RedPacket:
		return L"红包";
	case WeChatMsgType::Transfer:
		return L"转账";
	case WeChatMsgType::ChatRecord:
		return L"聊天记录";
	case WeChatMsgType::Location:
		return L"位置";
	case WeChatMsgType::VoiceCall:
		return L"音频通话";
	case WeChatMsgType::VideoCall:
		return L"视频通话";
	default:
		return L"";
	}
}

// 消息
struct WeChatMsg {
	WeChatMsgType m_type = WeChatMsgType::Unknown;  // 消息类型 
	std::wstring m_owner;     // 消息所有者
	std::wstring m_content;   // 消息内容
};

// 联系人
struct WeChatContact {
	std::wstring m_name;            // 联系人的名字
	std::wstring m_lastContactTime;        // 最近的联系时间
	std::wstring m_lastContactContent;     // 最近的联系内容
	std::vector<WeChatMsg> m_chatRecords; // 聊天记录
};

struct WeChatInfo {
	std::wstring m_loginUserName;           // 当前登陆的用户名
	std::vector<WeChatContact> m_contacts; // 最近联系人
	std::wstring m_curSelectedContact;      // 当前选中的联系人
};

void PrintDiff(const WeChatInfo& lastWeChat, const WeChatInfo& curWeChat);
void Merge(WeChatInfo& lastWeChat, WeChatInfo& curWeChat);

// 在主窗口中查找，包括：登录的用户名，最近联系人列表，当前正在聊天的联系人聊天记录
void FindInMainWnd(WeChatInfo* pWeChat, HWND hWnd);

// 在弹出的聊天窗口中查找，包括：聊天记录
void FindInChatWnd(WeChatInfo* pWeChat, HWND hWnd);

int main()
{
	HRESULT hr = CoInitialize(NULL);
	if (FAILED(hr)) {
		return 0;
	}

	bool bOpen = true;    // 是否已打开微信
	bool bLogin = true;   // 是否已登录
	WeChatInfo lastWeChat; // 存放上一次的查找结果，用于和本地的查找结果进行比较和合并
	while (true)
	{
		WeChatInfo curWeChat; // 存放本次的查找结果

		// 主窗口
		HWND hMainWnd = FindWindow(_T("WeChatMainWndForPC"), _T("微信"));  //获取微信窗口的句柄
		if (NULL == hMainWnd) {

			// 没找到主窗口，说明没有登陆或者没有打开微信
			HWND hLoginWnd = FindWindow(_T("WeChatLoginWndForPC"), _T("微信"));
			if (NULL == hLoginWnd) {
				// 没有找到登陆窗口，说明没有打开微信
				if (bOpen) { // 如果之前是打开状态，现在变为了非打开，则提示。防止重复提示
					std::cout << "请先打开微信客户端，并登陆微信" << std::endl;
					bOpen = false;
				}
			}
			else {
				// 找到了登陆窗口，提示用户登陆
				if (bLogin) { // 如果之前是登陆状态，现在变为了未登陆，则提示。防止重复提示
					std::cout << "您已打开微信客户端，请登陆微信" << std::endl;
					bLogin = false;
				}
			}	
		}
		else {
			if (!bLogin) {
				std::cout << "登陆成功" << std::endl;
				bLogin = true;
			}

			FindInMainWnd(&curWeChat, hMainWnd);
		}	

		// 所有弹出的聊天窗口
		HWND hChatWnd = NULL;
		while (true) {
			HWND hNextChatWnd = FindWindowEx(NULL, hChatWnd, _T("ChatWnd"), NULL);
			if (NULL == hNextChatWnd) {
				break;
			}
			FindInChatWnd(&curWeChat, hNextChatWnd);
			hChatWnd = hNextChatWnd;
		}

		// 比较本次的结果和上次的结果，输入变化
		PrintDiff(lastWeChat, curWeChat);

		// 将curWeChat 合并进lastWeChat中
		Merge(lastWeChat, curWeChat);

		Sleep(500);
	}

	CoUninitialize();
	return 0;
}

void PrintDiff(const WeChatInfo& lastWeChat, const WeChatInfo& curWeChat)
{
	std::wcout.imbue(std::locale("chs"));
	setlocale(LC_ALL, "");

	if (curWeChat.m_loginUserName != lastWeChat.m_loginUserName) {
		std::wcout << L"登陆用户名   ：" << curWeChat.m_loginUserName << std::endl;
	}

	if (curWeChat.m_curSelectedContact != lastWeChat.m_curSelectedContact) {
		std::wcout << L"当前选中联系人：" << curWeChat.m_curSelectedContact << std::endl;
	}

	for (int i = 0; i < curWeChat.m_contacts.size(); i++) {
		// 判断联系人是否已存在
		bool bExists = false;
		int  idx = 0;
		for (int j = 0; j < lastWeChat.m_contacts.size(); j++) {
			if (curWeChat.m_contacts[i].m_name == lastWeChat.m_contacts[j].m_name) {
				bExists = true;
				idx = j;
				break;
			}
		}

		// 已存在，则判断聊天记录是否有变化
		if (bExists) {
			// 如果大小有变化，先简单处理，只做增量
			size_t nLastCnt = lastWeChat.m_contacts[idx].m_chatRecords.size();
			size_t nCurCnt = curWeChat.m_contacts[i].m_chatRecords.size();

			if (nCurCnt > nLastCnt) {	
				wprintf(L"联系人：%-20s, 新增聊天记录：%zd 条\n", curWeChat.m_contacts[i].m_name.c_str(), nCurCnt - nLastCnt);

				for (int k = nLastCnt; k < nCurCnt; k++) {
					wprintf(L"     消息类型：%-10s, 消息发送者：%-20s，消息内容：%s\n",
						GetMsgTypeString(curWeChat.m_contacts[i].m_chatRecords[k].m_type).c_str(),
						curWeChat.m_contacts[i].m_chatRecords[k].m_owner.c_str(),
						curWeChat.m_contacts[i].m_chatRecords[k].m_content.c_str()
					);
				}
			}
		}

		// 不存在
		else {
			wprintf(L"联系人：%-20s, 最近联系时间：%-12s, 最近联系消息：%s\n",
				curWeChat.m_contacts[i].m_name.c_str(),
				curWeChat.m_contacts[i].m_lastContactTime.c_str(),
				curWeChat.m_contacts[i].m_lastContactContent.c_str()
			);

			for (int h = 0; h < curWeChat.m_contacts[i].m_chatRecords.size(); h++) {
				wprintf(L"     消息类型：%-10s, 消息发送者：%-20s，消息内容：%s\n",
					GetMsgTypeString(curWeChat.m_contacts[i].m_chatRecords[h].m_type).c_str(),
					curWeChat.m_contacts[i].m_chatRecords[h].m_owner.c_str(),
					curWeChat.m_contacts[i].m_chatRecords[h].m_content.c_str()
				);
			}
		}
	}
}

void Merge(WeChatInfo& lastWeChat, WeChatInfo& curWeChat)
{
	lastWeChat.m_loginUserName = curWeChat.m_loginUserName;
	lastWeChat.m_curSelectedContact = curWeChat.m_curSelectedContact;

	// 删除已经不存在的联系人
	/*for (auto it = lastWeChat.m_contacts.begin(); it != lastWeChat.m_contacts.end();) {

		bool bFind = false;
		for (auto itcur = curWeChat.m_contacts.begin(); itcur != curWeChat.m_contacts.end(); itcur++) {
			if (it->m_name == itcur->m_name) {
				bFind = true;
				break;
			}
		}

		if (!bFind) {
			it = lastWeChat.m_contacts.erase(it);
		}
		else{
		    it++;
		}
	}*/

	// 补充新增的联系人
	for (auto it = curWeChat.m_contacts.begin(); it != curWeChat.m_contacts.end(); it++) {
		
		bool bFind = false;
		for (auto itlast = lastWeChat.m_contacts.begin(); itlast != lastWeChat.m_contacts.end(); itlast++) {
			if (it->m_name == itlast->m_name) {
				bFind = true;

				itlast->m_lastContactTime = it->m_lastContactTime;
				itlast->m_lastContactContent = it->m_lastContactContent;

				// 如果联系人的聊天记录存在，并且比之前的要多，就直接覆盖
				if (it->m_chatRecords.size() > 0 && it->m_chatRecords.size() > itlast->m_chatRecords.size()) {
					itlast->m_chatRecords.swap(it->m_chatRecords);
				}
				break;
			}
		}

		if (!bFind) {
			// 未找到，则直接添加
			lastWeChat.m_contacts.push_back(*it);
		}
	}
}

// 从 e 中查找所有类型名称为 localizedControlType 的元素
IUIAutomationElementArray* FindElementArray(IUIAutomation* pAutomation, IUIAutomationElement* pRootElement, const WCHAR* localizedControlType, int* pArraySizeOut) {
	VARIANT v;
	v.vt = VT_BSTR;
	v.bstrVal = SysAllocString(localizedControlType);

	IUIAutomationCondition* pCondition;
	HRESULT hr = pAutomation->CreatePropertyCondition(UIA_LocalizedControlTypePropertyId, v, &pCondition);
	if (FAILED(hr)) {
		SysReleaseString(v.bstrVal);
		return NULL;
	}

	IUIAutomationElementArray* pArray;
	hr = pRootElement->FindAll(TreeScope_Descendants, pCondition, &pArray);
	if (FAILED(hr) || pArray == NULL) {
		pCondition->Release();
		SysReleaseString(v.bstrVal);
		return NULL;
	}

	pCondition->Release();
	SysReleaseString(v.bstrVal);

	hr = pArray->get_Length(pArraySizeOut);
	if (FAILED(hr)) {
		pArray->Release();
		return NULL;
	}

	return pArray;
}

std::wstring GetElementName(IUIAutomationElement* pElement)
{
	BSTR bName;
	HRESULT hr = pElement->get_CurrentName(&bName);
	if (FAILED(hr)) {
		return L"";
	}
	std::wstring wname(bName);
	SysFreeString(bName);

	return wname;
}

// 获取Array中元素的名称
std::wstring GetArrayItemName(IUIAutomationElementArray* pArray, int itemIndex) {
	
	int nArraySize = 0;
	HRESULT hr = pArray->get_Length(&nArraySize);
	if (FAILED(hr)) {
		return L"";
	}

	// 索引溢出
	if (itemIndex >= nArraySize) {
		return L"";
	}

	IUIAutomationElement* pElement;
	hr = pArray->GetElement(itemIndex, &pElement);
	if (FAILED(hr)) {
		return L"";
	}

	std::wstring name = GetElementName(pElement);
	pElement->Release();

	return name;
}

// 获取当前登陆的用户名
std::wstring FindWeChatUserName(IUIAutomation* pAutomation, IUIAutomationElement* pRootElement)
{
	// 找到所有的列表
	int nButtonCnt = 0;
	IUIAutomationElementArray* pButtonArray = FindElementArray(pAutomation, pRootElement, L"按钮", &nButtonCnt);
	if (NULL == pButtonArray) {
		return L"";
	}

	// 头像\聊天\通讯录 几个按钮是排在一起的，通过这个特征来找头像按钮，头像按钮的名称，就是登陆的用户名
	std::wstring name;
	for (int i = 0; i < nButtonCnt - 2; i++) {
		std::wstring btn1 = GetArrayItemName(pButtonArray, i + 1);
		std::wstring btn2 = GetArrayItemName(pButtonArray, i + 2);

		if (btn1 == L"聊天" && btn2 == L"通讯录") {
			name = GetArrayItemName(pButtonArray, i);
			break;
		}
	}
	pButtonArray->Release();
	return name;
}

// 处理 会话 列表，获取最近联系人列表
void FindRecentContactsInMainWnd(WeChatInfo* pWeChat, IUIAutomation* pAutomation, IUIAutomationElement* pListElement)
{
	int nArraySize = 0;
	IUIAutomationElementArray* pArray = FindElementArray(pAutomation, pListElement, L"列表项目", &nArraySize);
	if (NULL == pArray) return;

	// 遍历列表，即每个联系人信息
	for (int i = 0; i < nArraySize; i++) {
		IUIAutomationElement* pListItemElement;
		HRESULT hr = pArray->GetElement(i, &pListItemElement);
		if (FAILED(hr)) {
			continue;
		}

		// 如果该列表元素已经处理过，则忽略
		/*if (CheckExists(pAutomation, pListItemElement)) {
			pListItemElement->Release();
			continue;
		}*/

		WeChatContact contact;

		// 联系人名称
		contact.m_name = GetElementName(pListItemElement);

		// 当前是否被选中
		IUnknown * pProvider;
		hr = pListItemElement->GetCurrentPattern(UIA_SelectionItemPatternId, &pProvider);
		if (FAILED(hr)) {
			pListItemElement->Release();
			continue;
		}
		BOOL bIsSelected = 0;      // 当前是否被选中
		((ISelectionItemProvider*)pProvider)->get_IsSelected(&bIsSelected);

		if (bIsSelected) {
			pWeChat->m_curSelectedContact = contact.m_name;
		}
		
		// 从联系人 item 找到，最近的聊天时间和聊天内容
		int nTextArraySize = 0;
		IUIAutomationElementArray* pTextArray = FindElementArray(pAutomation, pListItemElement, L"文本", &nTextArraySize);
		if (pTextArray) {
			contact.m_lastContactTime = GetArrayItemName(pTextArray, 1);
			contact.m_lastContactContent = GetArrayItemName(pTextArray, 2);
			pTextArray->Release();
		}

		pWeChat->m_contacts.push_back(contact);
		pListItemElement->Release();
	}

	pArray->Release();
}

void ParseWeChatMsg(std::wstring loginUserName, WeChatMsg* pMsg, IUIAutomation* pAutomation, IUIAutomationElement* pElement) {
	
	// 获取所有按钮
	// 每条聊天记录都会有 按钮，按钮的名称就是消息发送者的名字
	// 如果没有按钮，则说明系统通知，或者是时间

	std::wstring content = GetElementName(pElement);

	int nButtonArraySize = 0;
	IUIAutomationElementArray* pButtonArray = FindElementArray(pAutomation, pElement, L"按钮", &nButtonArraySize);
	if (pButtonArray && nButtonArraySize > 0) {
		// 说明是聊天记录
		// 通过内容判断消息类型
		if (content == L"[图片]" || content == L"[Photo]") {
			pMsg->m_type = WeChatMsgType::Image;
		}
		else if (content == L"[视频]" || content == L"[Video]") {
			pMsg->m_type = WeChatMsgType::Video;
		}
		else if (content == L"[文件]" || content == L"[File]") {
			pMsg->m_type = WeChatMsgType::File;
		}
		else if (content == L"[链接]") {
			pMsg->m_type = WeChatMsgType::Link;
		}
		else if (content == L"[聊天记录]") {
			pMsg->m_type = WeChatMsgType::ChatRecord;
		}
		else if (content == L"[地理位置]") {
			pMsg->m_type = WeChatMsgType::Location;
		}
		else if (content == L"微信转账" && nButtonArraySize > 1) {
			pMsg->m_type = WeChatMsgType::Transfer;
		}
		else if (content.find(L"[语音]") != std::string::npos  && nButtonArraySize > 1) {
			pMsg->m_type = WeChatMsgType::Audio;
		}
		// todo 支持更多的格式，如视频通话/音频通话 等等
		else {
			pMsg->m_type = WeChatMsgType::Text;
		}

		pMsg->m_content = content;

		// 从头像按钮获取当前消息的Owner
		// 获取所有按钮的名称，查看是否与 当前登录用户名相同，如果有相同，说明这条消息是用户自己的消息
		// 否则，说明这条消息是好友发来的消息
		// 如果是好友的消息，那么头像按钮在最左边，即第一个 button
		bool isSelf = false;
		for (int i = 0; i < nButtonArraySize; i++) {
			std::wstring name = GetArrayItemName(pButtonArray, i);
			if (name == loginUserName) {
				isSelf = true;
				break;
			}
		}

		if (isSelf) {
			pMsg->m_owner = loginUserName;
		}
		else {
			pMsg->m_owner = GetArrayItemName(pButtonArray, 0); // 第一个button 就是好友的头像
		}
	}

	else {
		// 说明是系统通知或时间或者红包消息
		if (content.find(L':') != std::wstring::npos) {
			// 如果存在 : 说明是时间
			pMsg->m_type = WeChatMsgType::Time;
			pMsg->m_owner = L"-";
			pMsg->m_content = content;
		}

		else if (content == L"收到红包，请在手机上查看") {
			// 红包消息
			pMsg->m_type = WeChatMsgType::RedPacket;
			pMsg->m_owner = L"-";
			pMsg->m_content = content;
		}

		else {
			// 否则认为是系统通知
			pMsg->m_type = WeChatMsgType::Notify;
			pMsg->m_owner = L"-";
			pMsg->m_content = content;
		}
	}

	pButtonArray->Release();
}


// 处理 消息 列表，获取当前正在聊天的联系人的聊天记录
void FindCurrentChatRecordInMainWnd(WeChatInfo* pWeChat, IUIAutomation* pAutomation, IUIAutomationElement* pListElement)
{
	// 找到当前的联系人
	int contactIdx = 0;
	for (contactIdx = 0; contactIdx < pWeChat->m_contacts.size(); contactIdx++) {
		if (pWeChat->m_contacts[contactIdx].m_name == pWeChat->m_curSelectedContact) {
			break;
		}
	}

	// 当前没有选中联系人
	if (contactIdx >= pWeChat->m_contacts.size()) {
		return;
	}

	int nArraySize = 0;
	IUIAutomationElementArray* pArray = FindElementArray(pAutomation, pListElement, L"列表项目", &nArraySize);
	if (NULL == pArray) return;

	// 遍历列表，即每条聊天记录
	for (int i = 0; i < nArraySize; i++) {

		IUIAutomationElement* pListItemElement;
		HRESULT hr = pArray->GetElement(i, &pListItemElement);
		if (FAILED(hr)) {
			continue;
		}

		// 如果该列表元素已经处理过，则忽略
		/*if (CheckExists(pAutomation, pListItemElement)) {
			pListItemElement->Release();
			continue;
		}*/

		WeChatMsg msg;
		ParseWeChatMsg(pWeChat->m_loginUserName, &msg, pAutomation, pListItemElement);
		pWeChat->m_contacts[contactIdx].m_chatRecords.push_back(msg);

		pListItemElement->Release();
	}

	pArray->Release();

}

// 处理 消息 列表，从 ChatWnd 窗口中获取的聊天记录
// 注意：如果好友在单独的ChatWnd 时，它是不会显示在主窗口的联系人列表中的
// contactName: 当前聊天窗口的好友名称
void FindCurrentChatRecordInChatWnd(WeChatInfo* pWeChat, std::wstring contactName, IUIAutomation* pAutomation, IUIAutomationElement* pListElement)
{
	// 查看 contact 是否存在，正常情况下，不应该存在，因为如果是在单独窗口打开的话，就不会显示在主窗口
	for (auto it = pWeChat->m_contacts.begin(); it != pWeChat->m_contacts.end(); it++) {
		if (it->m_name == contactName) {
			// 如果已存在，说明除了问题，为了兼容，直接将该联系人数据删除即可
			pWeChat->m_contacts.erase(it);
			break;
		}
	}

	WeChatContact contact;
	contact.m_name = contactName;

	int nArraySize = 0;
	IUIAutomationElementArray* pArray = FindElementArray(pAutomation, pListElement, L"列表项目", &nArraySize);
	if (NULL == pArray) return;

	// 遍历列表，即每条聊天记录
	for (int i = 0; i < nArraySize; i++) {

		IUIAutomationElement* pListItemElement;
		HRESULT hr = pArray->GetElement(i, &pListItemElement);
		if (FAILED(hr)) {
			continue;
		}

		WeChatMsg msg;
		ParseWeChatMsg(pWeChat->m_loginUserName, &msg, pAutomation, pListItemElement);
		contact.m_chatRecords.push_back(msg);

		if (msg.m_type == WeChatMsgType::Time) {
			contact.m_lastContactTime = msg.m_content;
		}else {
			contact.m_lastContactContent = msg.m_content;
		}

		pListItemElement->Release();
	}

	pWeChat->m_contacts.push_back(contact);

	pArray->Release();

}


void FindInMainWnd(WeChatInfo* pWeChat, HWND hWnd)
{
	IUIAutomation* pAutomation;
	HRESULT hr = CoCreateInstance(CLSID_CUIAutomation, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pAutomation));
	if (FAILED(hr)) {
		return;
	}

	IUIAutomationElement* pRootElement = NULL;
	hr = pAutomation->ElementFromHandle(hWnd, &pRootElement);
	if (FAILED(hr) || pRootElement == NULL) {
		pAutomation->Release();
		return;
	}

	// 找到登录的用户名
	pWeChat->m_loginUserName = FindWeChatUserName(pAutomation, pRootElement);

	// 找到联系人列表和聊天记录
	// 联系人列表和聊天记录都是 列表类型，所以查找列表控件即可
	// 通常会有两个列表
	// 1. 左侧的“会话”列表，即联系人列表
	// 2. 右侧的“消息”列表，即聊天记录 (如果当前没有选择联系人，则不会有“消息”列表)

	int nListCnt = 0;
	IUIAutomationElementArray* pListArray = FindElementArray(pAutomation, pRootElement, L"列表", &nListCnt);
	if (NULL == pListArray) {
		return;
	}
	
	for (int i = 0; i < nListCnt; i++) {
		IUIAutomationElement* pListElement;
		hr = pListArray->GetElement(i, &pListElement);
		if (FAILED(hr)) {
			continue;
		}

		std::wstring name = GetElementName(pListElement);
		if (name == L"会话") {
			FindRecentContactsInMainWnd(pWeChat, pAutomation, pListElement);
		}else if (name == L"消息") {
			FindCurrentChatRecordInMainWnd(pWeChat, pAutomation, pListElement);
		}else {} // ignore

		pListElement->Release();
	}

	pListArray->Release();
	pRootElement->Release();
	pAutomation->Release();
	return;
}

void FindInChatWnd(WeChatInfo* pWeChat, HWND hWnd)
{
	IUIAutomation* pAutomation;
	HRESULT hr = CoCreateInstance(CLSID_CUIAutomation, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pAutomation));
	if (FAILED(hr)) {
		return;
	}

	IUIAutomationElement* pRootElement = NULL;
	hr = pAutomation->ElementFromHandle(hWnd, &pRootElement);
	if (FAILED(hr) || pRootElement == NULL) {
		pAutomation->Release();
		return;
	}

	// 找到当前的联系人名称
	std::wstring contactName = GetElementName(pRootElement);

	// 聊天记录
	// 聊天记录都是 列表类型，名称为 “消息”，所以查找列表控件即可
	int nListCnt = 0;
	IUIAutomationElementArray* pListArray = FindElementArray(pAutomation, pRootElement, L"列表", &nListCnt);
	if (NULL == pListArray) {
		return;
	}

	for (int i = 0; i < nListCnt; i++) {
		IUIAutomationElement* pListElement;
		hr = pListArray->GetElement(i, &pListElement);
		if (FAILED(hr)) {
			continue;
		}

		std::wstring name = GetElementName(pListElement);
		if (name == L"消息") {
			FindCurrentChatRecordInChatWnd(pWeChat, contactName, pAutomation, pListElement);
		}
		else {} // ignore

		pListElement->Release();
	}

	pListArray->Release();
	pRootElement->Release();
	pAutomation->Release();
	return;
}

