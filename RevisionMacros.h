#pragma once
#include <string>
#include <atlstr.h>
#include <atlimage.h>

#include "pch.h"

class RevisionMacros
{
private:
	_Application *spApp;
	BSTR Text;
	BSTR StyleName;

	float m_PortraitFirstMargin;
	float m_PortraitSecondMargin;
	float m_LandscapeFirstMargin;
	float m_LandscapeSecondMargin;

	void putPortraitFirstMargin(const float prop)
	{
		m_PortraitFirstMargin = spApp->MillimetersToPoints(prop);
	}
	
	float getPortraitFirstMargin()
	{
		return spApp->PointsToMillimeters(m_PortraitFirstMargin);
	}

	void putPortraitSecondMargin(const float prop)
	{
		m_PortraitSecondMargin = spApp->MillimetersToPoints(prop);
	}
	
	float getPortraitSecondMargin()
	{
		return spApp->PointsToMillimeters(m_PortraitSecondMargin);
	}

	void  putLandscapeFirstMargin(const float prop)
	{
		 m_LandscapeFirstMargin = spApp->MillimetersToPoints(prop);
	}
	
	float  getLandscapeFirstMargin()
	{
		return spApp->PointsToMillimeters(m_LandscapeFirstMargin);
	}

	void  putLandscapeSecondMargin(const float prop)
	{
		 m_LandscapeSecondMargin= spApp->MillimetersToPoints(prop);
	}
	
	float  getLandscapeSecondMargin()
	{
		return spApp->PointsToMillimeters( m_LandscapeSecondMargin);
	}

public:
	bool IsField = False;
	bool SilensStyleAppend = False;

	__declspec(property(
		get = getPortraitFirstMargin,
		put = putPortraitFirstMargin))
		float PortraitFirstMargin;
	__declspec(property(
		get = getPortraitSecondMargin,
		put = putPortraitSecondMargin))
		float PortraitSecondMargin;
	__declspec(property(
		get = getLandscapeFirstMargin,
		put = putLandscapeFirstMargin))
		float LandscapeFirstMargin;
	__declspec(property(
		get = getLandscapeSecondMargin,
		put = putLandscapeSecondMargin))
		float LandscapeSecondMargin;
	//__declspec(property(
	//	get = getStyleName,
	//	put = putStyleName))
	//	std::string StyleName;
	//__declspec(property(
	//	get = getBookmark,
	//	put = putBookmark))
	//	Word::Bookmark* Bookmark;


	RevisionMacros(_Application *App, BSTR Text)
	{

		this->spApp = App;
		this->Text = Text;

		this->PortraitFirstMargin = 0;
		this->PortraitSecondMargin = 20;
		this->LandscapeFirstMargin = 0;
		this->LandscapeSecondMargin = 5;
	}

	void Insert()
	{
		const auto doc = this->spApp->ActiveDocument;
		this->spApp->ScreenUpdating = False;

		BSTR m = SysAllocString(L"RevStyle");
		VARIANT revStyleName;
		revStyleName.bstrVal = m;
		revStyleName.vt = VT_BSTR;
		//CComPtr<Style> revStyle;
		Style *revStyle;
		auto r = doc->Styles->raw_Add(m, 0, &revStyle);
		if (r == S_OK)
		{
			revStyle->Font->Name = SysAllocString(L"Times New Roman");
			revStyle->AutomaticallyUpdate = False;
			revStyle->Font->Size = 12;
			revStyle->Font->Shading->Texture = wdTextureNone;
			revStyle->Font->Shading->ForegroundPatternColor = wdColorAutomatic;
			revStyle->Font->Shading->BackgroundPatternColor = wdColorAutomatic;
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

		auto firstPosition = rang->Information[wdVerticalPositionRelativeToPage];
		auto endPosition = erang->Information[wdVerticalPositionRelativeToPage];
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

