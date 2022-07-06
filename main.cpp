#include <stdio.h>
#include <string.h>
#include <iostream>
#include <tchar.h>
#include <windows.h>
#include <atlbase.h>
#include <comutil.h>  
#include <UIAutomation.h>
#include <vector>
#pragma comment(lib, "comsuppw.lib")
using namespace std;

void FindWechatRecord(HWND hWnd);
bool CheckExists(IUIAutomation* a, IUIAutomationElement* e);

std::vector<SAFEARRAY*> g_ids;

int _tmain(int argc, _TCHAR* argv[])
{
	HRESULT hr = CoInitialize(NULL);
	if (FAILED(hr)) {
		return 0;
	}

	while (true)
	{
		// 主窗口
		HWND hMainWnd = FindWindow(_T("WeChatMainWndForPC"), _T("微信"));  //获取微信窗口的句柄
		//HWND hMainWnd = FindWindow(_T("WeWorkWindow"), _T("企业微信"));		//获取企业微信窗口的句柄
		if (hMainWnd) {
			FindWechatRecord(hMainWnd);
		}

		// 所有弹出的聊天窗口
		HWND hChatWnd = NULL;
		while (true) {
			HWND hNextChatWnd = FindWindowEx(NULL, hChatWnd, _T("ChatWnd"), NULL);
			if (NULL == hNextChatWnd) {
				break;
			}

			FindWechatRecord(hNextChatWnd);
			hChatWnd = hNextChatWnd;
		}

		Sleep(2000);
	}

	CoUninitialize();
	return 0;
}

void FindWechatRecord(HWND hWnd)
{
	IUIAutomation* pAutomation;
	HRESULT hr = CoCreateInstance(CLSID_CUIAutomation, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pAutomation));
	if (FAILED(hr)) {
		cout << "22" << endl;
		return;
	}

	IUIAutomationElement* pRootElement = NULL;
	hr = pAutomation->ElementFromHandle(hWnd, &pRootElement);
	if (FAILED(hr) || pRootElement == NULL) {
		cout << "44" << endl;
		return;
	}

	WCHAR wcPaneName[] = L"列表项目";
	VARIANT trueVar;
	trueVar.vt = VT_BSTR;
	trueVar.bstrVal = SysAllocString(wcPaneName);

	IUIAutomationCondition* listItemPatternCondition;
	hr = pAutomation->CreatePropertyCondition(UIA_LocalizedControlTypePropertyId, trueVar, &listItemPatternCondition);  //设置条件接口
	if (FAILED(hr)) {
		cout << "55" << endl;
		return;
	}

	IUIAutomationElementArray* pListItemElements;
	hr = pRootElement->FindAll(TreeScope_Descendants, listItemPatternCondition, &pListItemElements);
	if (FAILED(hr) || pListItemElements == NULL) {
		cout << "66" << endl;
		return;
	}

	int listItemCount;
	hr = pListItemElements->get_Length(&listItemCount);
	if (FAILED(hr)) {
		return;
	}

	// 找到的列表元素包括 左侧的联系人列表 以及 右侧的聊天记录
	for (int i = 0; i < listItemCount; i++)
	{
		std::string log;
		IUIAutomationElement* pListItemElement;
		hr = pListItemElements->GetElement(i, &pListItemElement);
		if (FAILED(hr)) {
			return;
		}

		// 如果该列表元素已经处理过，则忽略
		if (CheckExists(pAutomation, pListItemElement)) {
			continue;
		}

		// 列表元素的名字
		// 在联系人列表时，就是联系人的名字
		// 在聊天记录元素中，就是聊天的内容
		BSTR bListItemElementName;
		hr = pListItemElement->get_CurrentName(&bListItemElementName);
		if (FAILED(hr)) {
			cout << "99" << endl;
			return;
		}

		CONTROLTYPEID ctid;
		pListItemElement->get_CachedControlType(&ctid);

		std::string listItemElementName = _com_util::ConvertBSTRToString(bListItemElementName);
		SysFreeString(bListItemElementName);
		log += listItemElementName;


		// 从列表元素中找“按钮元素”
		WCHAR PaneName[] = L"按钮";  //消息对象名称
		VARIANT Var;
		Var.vt = VT_BSTR;
		Var.bstrVal = SysAllocString(PaneName);

		IUIAutomationCondition* buttonPatternCondition;
		hr = pAutomation->CreatePropertyCondition(UIA_LocalizedControlTypePropertyId, Var, &buttonPatternCondition);
		if (FAILED(hr)) {
			return;
		}
		IUIAutomationElementArray* pButtonElements;
		pListItemElement->FindAll(TreeScope_Subtree, buttonPatternCondition, &pButtonElements);
		buttonPatternCondition->Release();

		int buttonElementCount = 0;
		hr = pButtonElements->get_Length(&buttonElementCount);
		if (FAILED(hr)) {
			return;
		}

		// 只处理最后一个按钮元素（其实，按钮元素，要么不存在，要么就是1个）
		if (buttonElementCount > 0) {
			IUIAutomationElement* pButtonElement;
			hr = pButtonElements->GetElement(buttonElementCount - 1, &pButtonElement);
			if (FAILED(hr)) {
				return;
			}

			BSTR bButtonElementName;
			hr = pButtonElement->get_CurrentName(&bButtonElementName);
			if (FAILED(hr)) {
				cout << "99" << endl;
				return;
			}

			string buttonElementName = _com_util::ConvertBSTRToString(bButtonElementName);
			SysFreeString(bButtonElementName);

			log += ":" + buttonElementName;
		}

		// 从列表元素中找 “标签元素”
		WCHAR labelPaneName[] = L"文本";  //消息对象名称
		VARIANT VarLabel;
		VarLabel.vt = VT_BSTR;
		VarLabel.bstrVal = SysAllocString(labelPaneName);

		IUIAutomationCondition* labelPatternCondition;
		hr = pAutomation->CreatePropertyCondition(UIA_LocalizedControlTypePropertyId, VarLabel, &labelPatternCondition);
		if (FAILED(hr)) {
			return;
		}
		IUIAutomationElementArray* pLabelElements;
		pListItemElement->FindAll(TreeScope_Subtree, labelPatternCondition, &pLabelElements);
		labelPatternCondition->Release();

		int labelElementCount = 0;
		hr = pLabelElements->get_Length(&labelElementCount);
		if (FAILED(hr)) {
			return;
		}

		// 处理所有的label
		for (int h = 0; h < labelElementCount; h++) {
			IUIAutomationElement* pLabelElement;
			hr = pLabelElements->GetElement(h, &pLabelElement);
			if (FAILED(hr)) {
				return;
			}

			BSTR bLabelElementName;
			hr = pLabelElement->get_CurrentName(&bLabelElementName);
			if (FAILED(hr)) {
				cout << "99" << endl;
				return;
			}

			string labelElementName = _com_util::ConvertBSTRToString(bLabelElementName);
			SysFreeString(bLabelElementName);

			log += ":" + labelElementName;
		}

		printf("%s\n", log.c_str());
	}

	listItemPatternCondition->Release();

	pAutomation->Release();

	std::cout << "......." << std::endl;
	return;
}

// 获取元素e的ID，判断该元素是否已经处理过
bool CheckExists(IUIAutomation* a, IUIAutomationElement* e)
{
	SAFEARRAY* pruntimeId;
	HRESULT hr = e->GetRuntimeId(&pruntimeId);
	for (int k = 0; k < g_ids.size(); k++) {
		BOOL bEqual = 0;
		a->CompareRuntimeIds(pruntimeId, g_ids[k], &bEqual);
		if (bEqual) return true;
	}

	// 不存在，则加入
	g_ids.push_back(pruntimeId);
	return false;
}
