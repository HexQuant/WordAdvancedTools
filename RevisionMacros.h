#pragma once
#include <string>
#include <atlstr.h>
#include <atlimage.h>

#include "pch.h"

class RevisionMacros
{
private:
	Word::_Application *spApp;

	float m_PortraitFirstMargin;
	float m_PortraitSecondMargin;
	float m_LandscapeFirstMargin;
	float m_LandscapeSecondMargin;

public:
	BSTR Text;
	BSTR StyleName;
	bool IsField = False;
	bool SilensStyleAppend = False;

	void putPortraitFirstMargin(const UINT prop)
	{
		m_PortraitFirstMargin = spApp->MillimetersToPoints(static_cast<float>(prop));
	}
	
	UINT getPortraitFirstMargin()
	{
		return static_cast<UINT>(spApp->PointsToMillimeters(m_PortraitFirstMargin));
	}

	void putPortraitSecondMargin(const UINT prop)
	{
		m_PortraitSecondMargin = spApp->MillimetersToPoints(static_cast<float>(prop));
	}
	
	UINT getPortraitSecondMargin()
	{
		return static_cast<UINT>(spApp->PointsToMillimeters(m_PortraitSecondMargin));
	}

	void putLandscapeFirstMargin(const UINT prop)
	{
		m_LandscapeFirstMargin = spApp->MillimetersToPoints(static_cast<float>(prop));
	}
	
	UINT getLandscapeFirstMargin()
	{
		return static_cast<UINT>(spApp->PointsToMillimeters(m_LandscapeFirstMargin));
	}

	void putLandscapeSecondMargin(const UINT prop)
	{
		m_LandscapeSecondMargin = spApp->MillimetersToPoints(static_cast<float>(prop));
	}
	
	UINT getLandscapeSecondMargin()
	{
		return static_cast<UINT>(spApp->PointsToMillimeters(m_LandscapeSecondMargin));
	}

	__declspec(property(
		get = getPortraitFirstMargin,
		put = putPortraitFirstMargin))
		UINT PortraitFirstMargin;
	__declspec(property(
		get = getPortraitSecondMargin,
		put = putPortraitSecondMargin))
		UINT PortraitSecondMargin;
	__declspec(property(
		get = getLandscapeFirstMargin,
		put = putLandscapeFirstMargin))
		UINT LandscapeFirstMargin;
	__declspec(property(
		get = getLandscapeSecondMargin,
		put = putLandscapeSecondMargin))
		UINT LandscapeSecondMargin;

	//__declspec(property(
	//	get = getStyleName,
	//	put = putStyleName))
	//	std::string StyleName;
	//__declspec(property(
	//	get = getBookmark,
	//	put = putBookmark))
	//	Word::Bookmark* Bookmark;


	RevisionMacros(Word::_Application *App)
	{

		this->spApp = App;

		Text = SysAllocString(L"1");
		StyleName = SysAllocString(L"RevStyle");

		this->PortraitFirstMargin = 0;
		this->PortraitSecondMargin = 20;
		this->LandscapeFirstMargin = 0;
		this->LandscapeSecondMargin = 5;
	}
	~RevisionMacros()
	{
		if (this->Text != NULL)
		{
			SysFreeString(this->Text);
		}
		if (this->StyleName != NULL)
		{
			SysFreeString(this->StyleName);
		}

	}

	void Insert()
	{
		const auto doc = this->spApp->ActiveDocument;
		this->spApp->ScreenUpdating = False;

		VARIANT revStyleName;
		revStyleName.bstrVal = this->StyleName;
		revStyleName.vt = VT_BSTR;
		Style *revStyle;
		auto r = doc->Styles->raw_Add(StyleName, 0, &revStyle);
		if (r == S_OK)
		{
			revStyle->Font->Name = SysAllocString(L"Times New Roman");
			revStyle->AutomaticallyUpdate = False;
			revStyle->Font->Size = 12;
			revStyle->Font->Shading->Texture = WdTextureIndex::wdTextureNone;
			revStyle->Font->Shading->ForegroundPatternColor = WdColor::wdColorAutomatic;
			revStyle->Font->Shading->BackgroundPatternColor = WdColor::wdColorAutomatic;
		}

		auto selection = this->spApp->Selection;
		auto rang = selection->Range;
		auto start = rang->Start;
		auto end = rang->End;
		VARIANT var1;
		var1.intVal = end - 1;
		var1.vt = VT_INT;
		VARIANT var2;
		var2.intVal = end;
		var2.vt = VT_INT;
		auto erang = doc->Range(&var1, &var2);
		auto pagesetup = rang->PageSetup;
		auto PageTopPosition = pagesetup->TopMargin;
		auto PageBottomMargin = pagesetup->BottomMargin;
		auto PageBottomPosition = pagesetup->PageHeight - PageBottomMargin;

		auto orientation = pagesetup->Orientation;
		auto papersize = pagesetup->PaperSize;

		float PortraitFirstMargin, PortraitSecondMargin,
			LandscapeFirstMargin, LandscapeSecondMargin;

		float width = 0, left = 0, heigth = 20;

		if ((orientation == wdOrientLandscape && papersize == wdPaperA3) ||
			(orientation == wdOrientPortrait && papersize == wdPaperA4))
		{
			left = this->m_PortraitFirstMargin;
			width = this->m_PortraitSecondMargin-left;
		}
		else
		{
			left = this->m_LandscapeFirstMargin;
			width = this->m_LandscapeSecondMargin - left;
		}

		auto firstPosition = rang->Information[WdInformation::wdVerticalPositionRelativeToPage];
		auto endPosition = erang->Information[WdInformation::wdVerticalPositionRelativeToPage];
		auto endCharHeight = erang->Font->Size * 1.152083417;

		heigth = endPosition.fltVal + endCharHeight - firstPosition.fltVal;
		if (heigth <= 0)
		{
			return;
		}

		auto shape = doc->Shapes->AddTextbox(msoTextOrientationHorizontal, left, firstPosition.fltVal, width, heigth, NULL);
		shape->Line->Visible = msoFalse;

		auto textFrame = shape->TextFrame;

		textFrame->MarginTop = 0;
		textFrame->MarginBottom = 0;
		textFrame->MarginLeft = 0;
		textFrame->MarginRight = 0;

		auto textRange = textFrame->TextRange;

		textRange->PutStyle(&revStyleName);

		textRange->Text = this->Text;

		textRange->ParagraphFormat->FirstLineIndent = 0;
		textRange->ParagraphFormat->Alignment = wdAlignParagraphRight;
		textRange->Borders->Item(wdBorderTop)->LineStyle = wdLineStyleNone;
		textRange->Borders->Item(wdBorderLeft)->LineStyle = wdLineStyleNone;
		textRange->Borders->Item(wdBorderBottom)->LineStyle = wdLineStyleNone;
		textRange->Borders->Item(wdBorderRight)->LineStyle = this->spApp->Options->DefaultBorderLineStyle;
		textRange->Borders->Item(wdBorderRight)->LineWidth = this->spApp->Options->DefaultBorderLineWidth;


		this->spApp->ScreenUpdating = True;
		this->spApp->ScreenRefresh();


	}

};

