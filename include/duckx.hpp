/*
 * Under MIT license
 * Author: Amir Mohamadi (@amiremohamadi)
 * DuckX is a free library to work with docx files.
 */
#pragma once
#ifndef DUCKX_H
#define DUCKX_H

//#define DUCKX_EXPORT __declspec(dllexport)
#ifdef DUCKX_EXPORTS
    #define DUCKX_EXPORT __declspec(dllexport)
#else
    #define DUCKX_EXPORT __declspec(dllimport)
#endif

#include <cstdio>
#include <stdlib.h>
#include <string>
#include <vector>
#include <utility>

#include <constants.hpp>
#include <duckxiterator.hpp>
#include "pugixml/pugixml.hpp"
#include "zip/zip.h"

// TODO: Use container-iterator design pattern!

 namespace duckx {

    enum class align
    {
        left,
        middle,
        right
    };

    enum class styles
    {
        style,
        table,
        column,
        row,
        cell,
        paragraph,
        run
    };

// Run contains runs in a paragraph
class DUCKX_EXPORT Run {
  private:
    friend class IteratorHelper;
    // Store the parent node (a paragraph)
    pugi::xml_node parent;
    // And store current node also
    pugi::xml_node current;

  public:
    Run();
    Run(pugi::xml_node, pugi::xml_node);
    void set_parent(pugi::xml_node);
    void set_current(pugi::xml_node);

    std::string get_text() const;
    bool set_text(const std::string &) const;
    bool set_text(const char *) const;

    Run &next();
    bool has_next() const;
};

// Paragraph contains a paragraph
// and stores runs
class DUCKX_EXPORT Paragraph {
  private:
    friend class IteratorHelper;
    // Store parent node (usually the body node)
    pugi::xml_node parent;
    // And store current node also
    pugi::xml_node current;
    // A paragraph consists of runs
    Run run;

  public:
    Paragraph();
    Paragraph(pugi::xml_node, pugi::xml_node);
    void set_parent(pugi::xml_node);
    void set_current(pugi::xml_node);
    // DELETE
    void delete_par() { parent.remove_child(current); }
    Paragraph &next();
    bool has_next() const;

    Run &runs();
    Run &add_run(const std::string &, std::string stylename = "RegText");
    Run &add_run(const char *, const char* stylename = "RegText");
    Paragraph &insert_paragraph_after(const std::string &,
        std::string stylename = "P1");

    void add_image(const std::string& name, const std::string& width = "", const std::string& height = "");
    void set_style(const std::string& name);
};

// TableCell contains one or more paragraphs
class DUCKX_EXPORT TableCell {
  private:
    friend class IteratorHelper;
    pugi::xml_node parent;
    pugi::xml_node current;

    Paragraph paragraph;

  public:
    TableCell();
    TableCell(pugi::xml_node, pugi::xml_node);

    void set_parent(pugi::xml_node);
    void set_current(pugi::xml_node);

    Paragraph &paragraphs();

    TableCell &next();
    bool has_next() const;
    Paragraph& add_paragraph(const std::string&);
};

// TableRow consists of one or more TableCells
class DUCKX_EXPORT TableRow {
    friend class IteratorHelper;
    pugi::xml_node parent;
    pugi::xml_node current;

    TableCell cell;

  public:
    TableRow();
    TableRow(pugi::xml_node, pugi::xml_node);
    void set_parent(pugi::xml_node);
    void set_current(pugi::xml_node);
    void delete_row() { parent.remove_child(current); }
    TableCell &cells();
    TableCell& add_cell(const std::string& cellstyle, const std::string& parstyle);
    TableCell& add_cell(const std::string& cellstyle);
    void add_covered_cell();
    TableCell& add_united_cell(const std::string& cellstyle, const std::string& parstyle, const int united_cell_columns, const int united_cell_rows = 1);
    bool has_next() const;
    TableRow &next();
};

// Table consists of one or more TableRow objects
class DUCKX_EXPORT Table {
  private:
    friend class IteratorHelper;
    pugi::xml_node parent;
    pugi::xml_node current;

    TableRow row;

  public:
    Table();
    Table(pugi::xml_node, pugi::xml_node);
    void set_parent(pugi::xml_node);
    void set_current(pugi::xml_node);

    Table &next();
    bool has_next() const;

    TableRow &rows();
    TableRow& add_row(const std::string& stylename);
    void add_column(const std::vector<std::string>& stylenames);
};
class DUCKX_EXPORT Style
{

private:
    friend class IteratorHelper;
    pugi::xml_node parent;
    pugi::xml_node current;
public:
    Style();
    Style(pugi::xml_node, pugi::xml_node);
    void set_parent(pugi::xml_node);
    void set_current(pugi::xml_node);

    Style& next();
    bool has_next() const;

    Style& add_style(std::string stylename, styles st, std::vector<std::pair<std::string, std::string>> attr);

};

// Document contains whole the docx file
// and stores paragraphs
class DUCKX_EXPORT Document {
  private:
    friend class IteratorHelper;
    std::string directory;
    Paragraph paragraph;
    Table table;
    Style style;
    pugi::xml_document document;

  public:
    Document();
    Document(std::string);
    void file(std::string);
    void open();
    void save() const;
    void save_copy(std::string) const;

    Paragraph &paragraphs();
    Table &tables();
    Style& styles();

    Table& add_table(const std::string& stylename);
    Paragraph& add_paragraph(const std::string& stylename);


};
} // namespace duckx

#endif
