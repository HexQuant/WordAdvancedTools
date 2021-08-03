#pragma once

#include "pch.h"

class BookmarkMgr
{
private:
	Word::_Application *spApp;
	ATL::CComPtr<Word::Bookmark> m_Bookmark;
	bool m_IsPined = false;
public:
	BookmarkMgr(Word::_Application * App)
	{
		this->spApp = App;
	}

	bool PinBookmark()
	{
		const auto doc = this->spApp->ActiveDocument;
		auto bookmark_count = spApp->Selection->Bookmarks->Count;
		if (bookmark_count == 1)
		{
			VARIANT i;
			i.uintVal = 1;
			i.vt = VT_UINT;
			m_Bookmark = spApp->Selection->Bookmarks->Item(&i);
			m_IsPined = true;
			return true;
		}
		else if (bookmark_count == 0)
		{
			MessageBoxW(NULL, L"No bookmarks found in selection", L"Bookmark", MB_OK | MB_ICONWARNING);
			return false;
		}
		else
		{
			MessageBoxW(NULL, L"There is more than one bookmark in the selection", L"Bookmark", MB_OK | MB_ICONWARNING);
			return false;
		}

		return true;
	}

	bool AddBookmarksBySelText()
	{
		auto sel = this->spApp->Selection;
		auto parag_count = sel->Paragraphs->Count;
		for (UINT i = 1; i <= parag_count; i++)
		{
			auto text = 3;

		}
		return true;
	}
	bool InsertRefByText()
	{
		return true;
	}
	bool InsertRefByID()
	{
		return true;

	}
	void GetSelBookmarksID()
	{

	}

};

