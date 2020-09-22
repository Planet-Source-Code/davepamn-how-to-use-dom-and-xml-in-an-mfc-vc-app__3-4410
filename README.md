<div align="center">

## How to Use DOM and XML in an MFC VC\+\+ App


</div>

### Description

Overview: This article explains how to program Visual C++ to use DOM for XML manipulation.

XML is composed of elements.  An element is a user defined tag name with a opening and close pair: such as, <Name> </Name>  Attributes are name/value pair combinations within an element. <Name first_name=David last_name=Nishimoto> </Name>

The Document Object Model (DOM) is used to add, update, and remove elements and attributes. This

article will explain how to manipulate xml elements and attributes using Visual C++ and DOM.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[davepamn](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/davepamn.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |Microsoft Visual C\+\+
**Category**       |[Algorithms](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/algorithms__3-29.md)
**World**          |[C / C\+\+](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/c-c.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/davepamn-how-to-use-dom-and-xml-in-an-mfc-vc-app__3-4410/archive/master.zip)





### Source Code

<html><head><title>Visual C++ and XML</title></head>
<body>
<pre>
Author: David Nishimoto
email: davepamn@relia.net
<a href="http://www.listensoftware.com/xmldom.zip">Download Sample</a>
Overview: This article explains how to program Visual C++ to use DOM and
xml.
XML is composed of elements.  An element is a user defined tag name with
a opening and close pair: such as, &lt;Name&gt; &lt;/Name&gt;  Attributes are name/value
pair combinations within an element. &lt;Name first_name=David last_name=Nishimoto&gt; &lt;/Name&gt;
The Document Object Model (DOM) is used to add, update, and remove elements and attributes. This
article will explain how to manipulate xml elements and attributes using Visual C++.
</pre>
<h3>CXMLEngine Class xmlengine.h</h3>
<pre>
1. The DOM object is a pointer which the implementation instantiates
2. The XML character set support unicode code points
class CXMLEngine
{
public:
	IXMLDOMDocumentPtr objDOMDoc;
	BSTR mXML;
public:
	CXMLEngine();
	~CXMLEngine();
	int CXMLEngine::Initialize(char *sFileName);
	IXMLDOMNodePtr GetNode(BSTR sKey);
	HRESULT DeleteNode(BSTR sKey);
	VARIANT GetAttribute(IXMLDOMNodePtr oNode,char *sKey);
	HRESULT SaveXML_To_File(char *filename);
	HRESULT AddNode(char *sParent_Key, char *sKey, char *sElementInformation);
	//HRESULT SetAttribute(IXMLDOMElement *oNode,char *sAttributeName, char *sAttributeValue);
	void Refresh();
};
</pre>
<h3>The DOM COM object is instantiated xmlengine.cpp</h3>
<pre>
1. The DOM COM object instance is instantiated
2. A XML file is loaded into the DOM object
int CXMLEngine::Initialize(char *sFileName)
{
	char searchPath[200];
	try
  {
		EVAL_HR(CoInitialize(NULL));
		//EVAL_HR(objDOMDoc.CreateInstance("Msxml2.DOMDocument.3.0"));
		EVAL_HR(objDOMDoc.CreateInstance("microsoft.xmldom"));
		GetCurrentDirectory(200, searchPath);
		strcat(searchPath,"\\");
		strcat(searchPath,sFileName);
		_variant_t varXml(searchPath);
		_variant_t varOut((bool)TRUE);
		objDOMDoc->async = false;
		varOut = objDOMDoc->load(varXml);
		mXML=objDOMDoc->xml;
		if ((bool)varOut == FALSE)
		 throw(0);
		return 0;
	}
	catch(...)
	{
		AfxMessageBox("Exception occurred");
		return -1;
	}
	CoUninitialize();
}
</pre>
<h3>Element Manipulation</h3>
<pre>
IXMLDOMNodePtr CXMLEngine::GetNode(BSTR sKey)
{
/*
Purpose:locate a specific node in the
xml structure by its "key" value
*/
	CString sCriteria;
  IXMLDOMNodePtr oNode;
  sCriteria = "//node[@key $eq$ '";
	sCriteria += sKey;
	sCriteria += "']";
	oNode=objDOMDoc->selectSingleNode(_bstr_t(sCriteria));
  if (oNode!=NULL)
	{
    return(oNode);
	}
	else
	{
		AfxMessageBox("Node Not Found");
		return(NULL);
	}
}
HRESULT CXMLEngine::DeleteNode(BSTR sKey)
{
/*
'Purpose: remove a node from the xml tree
'1)Find the node
'2)Find the node's parent
'3)Invoke the parents removeChild method deleting
'the node and its descendants from the xml tree.
*/
  IXMLDOMNodePtr oNode;
  IXMLDOMNodePtr oParentNode;
  VARIANT sParent_Key;
  oNode = GetNode(sKey);
  if(oNode != NULL)
	{
		sParent_Key=GetAttribute(oNode,"parent_key");
		//AfxMessageBox(_bstr_t(sParent_Key));
    oParentNode = GetNode(_bstr_t(sParent_Key));
    if (oParentNode != NULL)
		{
      oParentNode->removeChild(oNode);
		}
	}
	else
	{
		return(E_FAIL);
	}
	return(S_OK);
}
HRESULT CXMLEngine::AddNode(char *sParent_Key, char *sKey, char *sElementInformation)
{
/*
Purpose: Add a node to the xml tree
1)Search for the parent node
2)Assign attribute values to the new node
3)Insert the new node into the xml tree
*/
  //IXMLDOMNodePtr oNewNode;
	IXMLDOMElementPtr oElementNode;
  IXMLDOMNodePtr oParentNode, oProposedNode;
	CString sCriteria;
  HRESULT hr;
	//1-Check if the Node already exists
  sCriteria = "//node[@key $eq$ '";
	sCriteria += sKey;
	sCriteria += "']";
  oProposedNode = objDOMDoc->selectSingleNode(_bstr_t(sCriteria));
	if (oProposedNode!=NULL)
	{
		return (E_FAIL);
	}
  //2-Search for Parent Node Match
  sCriteria = "//node[@key $eq$ '";
	sCriteria += sParent_Key;
	sCriteria += "']";
  oParentNode = objDOMDoc->selectSingleNode(_bstr_t(sCriteria));
  //3-Insert new Node into Document Object Model (dom) structure
  if (oParentNode != NULL)
	{
	  //oNewNode = objDOMDoc->createElement("node");
		oElementNode=objDOMDoc->createElement("node");
		hr=oElementNode->setAttribute(_bstr_t("parent_key"),_bstr_t(sParent_Key));
		hr=oElementNode->setAttribute(_bstr_t("key"),_bstr_t(sKey));
 		hr=oElementNode->setAttribute(_bstr_t("element_inforamtion"),_bstr_t(sElementInformation));
		oParentNode->appendChild(oElementNode);
  }
	else
	{
		return(E_FAIL);
	}
	return(S_OK);
}
</pre>
<h3>Attribute Manipulation</h3>
<pre>
VARIANT CXMLEngine::GetAttribute(IXMLDOMNodePtr oNode,char *sKey)
{
	IXMLDOMNodePtr oAttribute;
	IXMLDOMNamedNodeMapPtr DOMNamedNodeMapPtr;
	VARIANT sReturn_Key;
	DOMNamedNodeMapPtr = oNode->attributes;
	oAttribute=DOMNamedNodeMapPtr->getNamedItem(_bstr_t(sKey));
	oAttribute->get_nodeValue(&sReturn_Key);
	return(sReturn_Key);
}
Setting an Attribute inside of an element
	hr=oElementNode->setAttribute(_bstr_t("parent_key"),_bstr_t(sParent_Key));
	hr=oElementNode->setAttribute(_bstr_t("key"),_bstr_t(sKey));
 	hr=oElementNode->setAttribute(_bstr_t("element_inforamtion"),_bstr_t(sElementInformation));
</pre>
</body>
</html>

