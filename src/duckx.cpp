#include "duckx.hpp"
#include <cctype>

// Hack on pugixml
// We need to write xml to std string (or char *)
// So overload the write function
struct xml_string_writer : pugi::xml_writer {
    std::string result;

    virtual void write(const void *data, size_t size) {
        result.append(static_cast<const char *>(data), size);
    }
};

duckx::Run::Run() {}

duckx::Run::Run(pugi::xml_node parent, pugi::xml_node current) {
    this->set_parent(parent);
    this->set_current(current);
}

void duckx::Run::set_parent(pugi::xml_node node) {
    this->parent = node;
    this->current = this->parent.child("text:span");
}

void duckx::Run::set_current(pugi::xml_node node) { this->current = node; }

std::string duckx::Run::get_text() const {
    return this->current.text().get();
}

bool duckx::Run::set_text(const std::string &text) const {
    return this->current.text().set(text.c_str());
}

bool duckx::Run::set_text(const char *text) const {
    return this->current.text().set(text);
}

duckx::Run &duckx::Run::next() {
    this->current = this->current.next_sibling("text:span");
    return *this;
}

bool duckx::Run::has_next() const { return this->current != 0; }

// Table cells
duckx::TableCell::TableCell() {}

duckx::TableCell::TableCell(pugi::xml_node parent, pugi::xml_node current) {
    this->set_parent(parent);
    this->set_current(current);
}

void duckx::TableCell::set_parent(pugi::xml_node node) {
    this->parent = node;
    this->current = this->parent.child("table:table-cell");
    this->paragraph.set_parent(this->current);
}

void duckx::TableCell::set_current(pugi::xml_node node) {
    this->current = node;
}

bool duckx::TableCell::has_next() const { return this->current != 0; }

duckx::Paragraph& duckx::TableCell::add_paragraph(const std::string& text)
{
    pugi::xml_node new_para =
        this->current.append_child("text:p");

    Paragraph* p = new Paragraph();
    p->set_current(new_para);
    if (text.size() != 0)
        p->add_run(text);

    return *p;
}

duckx::TableCell &duckx::TableCell::next() {
    this->current = this->current.next_sibling("table:table-cell");
    return *this;
}

duckx::Paragraph &duckx::TableCell::paragraphs() {
    this->paragraph.set_parent(this->current);
    return this->paragraph;
}

// Table rows
duckx::TableRow::TableRow() {}

duckx::TableRow::TableRow(pugi::xml_node parent, pugi::xml_node current) {
    this->set_parent(parent);
    this->set_current(current);
}

void duckx::TableRow::set_parent(pugi::xml_node node) {
    this->parent = node;
    this->current = this->parent.child("table:table-row");
    this->cell.set_parent(this->current);
}

void duckx::TableRow::set_current(pugi::xml_node node) { this->current = node; }

duckx::TableRow &duckx::TableRow::next() {
    this->current = this->current.next_sibling("table:table-row");
    return *this;
}

duckx::TableCell &duckx::TableRow::cells() {
    this->cell.set_parent(this->current);
    return this->cell;
}

duckx::TableCell& duckx::TableRow::add_cell(const std::string& cellstyle, const std::string& parstyle)
{
    // Add new run
    pugi::xml_node new_cell = this->current.append_child("table:table-cell");
    new_cell.append_attribute("table:style-name").set_value(cellstyle.c_str());
    new_cell.append_child("text:p").append_attribute("text:style-name").set_value(parstyle.c_str());

    return *new TableCell(this->current, new_cell);
}

duckx::TableCell& duckx::TableRow::add_cell(const std::string& cellstyle)
{
    // Add new run
    pugi::xml_node new_cell = this->current.append_child("table:table-cell");
    new_cell.append_attribute("table:style-name").set_value(cellstyle.c_str());

    return *new TableCell(this->current, new_cell);
}

void duckx::TableRow::add_covered_cell()
{
    //pugi::xml_node new_cell = 
        this->current.append_child("table:covered-table-cell");
    //return *new TableCell(this->current, new_cell);
}

duckx::TableCell& duckx::TableRow::add_united_cell(const std::string& cellstyle, const std::string& parstyle, const int united_cell_columns, const int united_cell_rows)
{
    pugi::xml_node new_cell = this->current.append_child("table:table-cell");
    new_cell.append_attribute("table:style-name").set_value(cellstyle.c_str());
    new_cell.append_attribute("table:number-columns-spanned").set_value(united_cell_columns);
    if (united_cell_rows > 1)
        new_cell.append_attribute("table:number-rows-spanned").set_value(united_cell_rows);
    new_cell.append_child("text:p").append_attribute("text:style-name").set_value(parstyle.c_str());
    for (int i = 1; i < united_cell_columns; i++)
        this->current.append_child("table:covered-table-cell");
    
    return *new TableCell(this->current, new_cell);
}

bool duckx::TableRow::has_next() const { return this->current != 0; }

// Tables
duckx::Table::Table() {}

duckx::Table::Table(pugi::xml_node parent, pugi::xml_node current) {
    this->set_parent(parent);
    this->set_current(current);
}

void duckx::Table::set_parent(pugi::xml_node node) {
    this->parent = node;
    this->current = this->parent.child("table:table");
    this->row.set_parent(this->current);
}

bool duckx::Table::has_next() const { return this->current != 0; }

duckx::Table &duckx::Table::next() {
    this->current = this->current.next_sibling("table:table");
    this->row.set_parent(this->current);
    return *this;
}

void duckx::Table::set_current(pugi::xml_node node) { this->current = node; }

duckx::TableRow &duckx::Table::rows() {
    this->row.set_parent(this->current);
    return this->row;
}

duckx::TableRow& duckx::Table::add_row(const std::string& stylename)
{
    // Add new run
    pugi::xml_node new_row = this->current.append_child("table:table-row");
    new_row.append_attribute("table:style-name").set_value(stylename.c_str());

    return *new TableRow(this->current, new_row);
}

void duckx::Table::add_column(const std::vector<std::string>& stylenames)
{
    pugi::xml_node new_cols = this->current.append_child("table:table-columns");
    for (auto elem : stylenames)
        new_cols.append_child("table:table-column").append_attribute("table:style-name").set_value(elem.c_str());
}


duckx::Paragraph::Paragraph() {}

duckx::Paragraph::Paragraph(pugi::xml_node parent, pugi::xml_node current) {
    this->set_parent(parent);
    this->set_current(current);
}

void duckx::Paragraph::set_parent(pugi::xml_node node) {
    this->parent = node;
    this->current = this->parent.child("text:p");

    this->run.set_parent(this->current);
}

void duckx::Paragraph::set_current(pugi::xml_node node) {
    this->current = node;
}

duckx::Paragraph &duckx::Paragraph::next() {
    this->current = this->current.next_sibling("text:p");
    this->run.set_parent(this->current);
    return *this;
}

bool duckx::Paragraph::has_next() const { return this->current != 0; }

duckx::Run &duckx::Paragraph::runs() {
    this->run.set_parent(this->current);
    return this->run;
}

duckx::Run &duckx::Paragraph::add_run(const std::string &text,
    std::string stylename) {
    return this->add_run(text.c_str(), stylename.c_str());
}

duckx::Run &duckx::Paragraph::add_run(const char *text,
    const char* stylename) {
    // Add new run
    pugi::xml_node new_run = this->current.append_child("text:span");
    //// Insert meta to new run
    new_run.append_attribute("text:style-name").set_value(stylename);
    
    // If the run starts or ends with whitespace characters, preserve them using
    // the xml:space attribute
    //if (*text != 0 && (isspace(text[0]) || isspace(text[strlen(text) - 1])))
    //    new_run.append_attribute("xml:space").set_value("preserve");
    //new_run_text.text().set(text);
    new_run.text().set(text);
    return *new Run(this->current, new_run);
}

duckx::Paragraph &
duckx::Paragraph::insert_paragraph_after(const std::string &text,
                                         std::string stylename) {

    pugi::xml_node new_para =
        this->parent.insert_child_after("text:p", this->current);

    Paragraph *p = new Paragraph();
    p->set_current(new_para);
    if (text.size() != 0)
        p->add_run(text, stylename);

    return *p;
}

void duckx::Paragraph::add_image(const std::string& name, const std::string& width/* = ""*/, const std::string& height/* = ""*/)
{
    pugi::xml_node new_frame =
        this->current.append_child("draw:frame");
    //new_frame.append_attribute("draw:style-name").set_value("a0");
    new_frame.append_attribute("text:anchor-type").set_value("as-char");
    if (width.size() > 0)
    {
        new_frame.append_attribute("svg:width").set_value(width.c_str());
        if (height.size() > 0)
            new_frame.append_attribute("svg:height").set_value(height.c_str());
        else
            new_frame.append_attribute("svg:height").set_value(width.c_str());
    }
    else
    {
        new_frame.append_attribute("svg:width").set_value("2.70833in");
        new_frame.append_attribute("svg:height").set_value("1.35833in");
    }
    new_frame.append_attribute("style:rel-width").set_value("scale");
    new_frame.append_attribute("style:rel-height").set_value("scale");

    pugi::xml_node new_image = new_frame.append_child("draw:image");
    new_image.append_attribute("xlink:href").set_value(std::string("media/").append(name).c_str());


}

void duckx::Paragraph::set_style(const std::string& name)
{
    current.attribute("text:style-name").set_value(name.c_str());
}

duckx::Document::Document() {
    // TODO: this function must be removed!
    this->directory = "";
}

duckx::Document::Document(std::string directory) {
    this->directory = directory;
}

void duckx::Document::file(std::string directory) {
    this->directory = directory;
}

void duckx::Document::open() {
    void *buf = NULL;
    size_t bufsize;

    // Open file and load "xml" content to the document variable
    zip_t *zip =
        zip_open(this->directory.c_str(), ZIP_DEFAULT_COMPRESSION_LEVEL, 'r');

    //zip_entry_open(zip, "word/document.xml");
    zip_entry_open(zip, "content.xml");
    zip_entry_read(zip, &buf, &bufsize);

    zip_entry_close(zip);
    zip_close(zip);

    this->document.load_buffer(buf, bufsize);

    free(buf);

    //this->paragraph.set_parent(document.child("w:document").child("w:body"));
    this->paragraph.set_parent(document.child("office::document-content").child("office:body").child("office:text"));
}

void duckx::Document::save() const {
    // minizip only supports appending or writing to new files
    // so we must
    // - make a new file
    // - write any new files
    // - copy the old files
    // - delete old docx
    // - rename new file to old file

    // Read document buffer
    xml_string_writer writer;
    this->document.print(writer);

    // Open file and replace "xml" content

    std::string original_file = this->directory;
    std::string temp_file = this->directory + ".tmp";

    // Create the new file
    zip_t *new_zip =
        zip_open(temp_file.c_str(), ZIP_DEFAULT_COMPRESSION_LEVEL, 'w');

    // Write out document.xml
    //zip_entry_open(new_zip, "word/document.xml");
    zip_entry_open(new_zip, "content.xml");
    const char *buf = writer.result.c_str();

    zip_entry_write(new_zip, buf, strlen(buf));
    zip_entry_close(new_zip);

    // Open the original zip and copy all files which are not replaced by duckX
    zip_t *orig_zip =
        zip_open(original_file.c_str(), ZIP_DEFAULT_COMPRESSION_LEVEL, 'r');

    // Loop & copy each relevant entry in the original zip
    int orig_zip_entry_ct = zip_total_entries(orig_zip);
    for (int i = 0; i < orig_zip_entry_ct; i++) {
        zip_entry_openbyindex(orig_zip, i);
        const char *name = zip_entry_name(orig_zip);

        // Skip copying the original file
        //if (std::string(name) != std::string("word/document.xml")) {
        if (std::string(name) != std::string("media/"))
        if (std::string(name) != std::string("content.xml")) {
            // Read the old content
            void *entry_buf;
            size_t entry_buf_size;
            zip_entry_read(orig_zip, &entry_buf, &entry_buf_size);

            // Write into new zip
            zip_entry_open(new_zip, name);
            zip_entry_write(new_zip, entry_buf, entry_buf_size);
            zip_entry_close(new_zip);

            free(entry_buf);
        }

        zip_entry_close(orig_zip);
    }

    // Close both zips
    zip_close(orig_zip);
    zip_close(new_zip);

    // Remove original zip, rename new to correct name
    remove(original_file.c_str());
    rename(temp_file.c_str(), original_file.c_str());
}

void duckx::Document::save_copy(std::string new_name) const {
    // minizip only supports appending or writing to new files
    // so we must
    // - make a new file
    // - write any new files
    // - copy the old files
    // - delete old docx
    // - rename new file to old file

    // Read document buffer
    xml_string_writer writer;
    this->document.print(writer);

    // Open file and replace "xml" content

    std::string original_file = this->directory;
    std::string temp_file = this->directory + ".tmp";

    // Create the new file
    zip_t* new_zip =
        zip_open(temp_file.c_str(), ZIP_DEFAULT_COMPRESSION_LEVEL, 'w');

    // Write out document.xml
    zip_entry_open(new_zip, "content.xml");
    const char* buf = writer.result.c_str();

    zip_entry_write(new_zip, buf, strlen(buf));
    zip_entry_close(new_zip);

    // Open the original zip and copy all files which are not replaced by duckX
    zip_t* orig_zip =
        zip_open(original_file.c_str(), ZIP_DEFAULT_COMPRESSION_LEVEL, 'r');

    // Loop & copy each relevant entry in the original zip
    int orig_zip_entry_ct = zip_total_entries(orig_zip);
    for (int i = 0; i < orig_zip_entry_ct; i++) {
        zip_entry_openbyindex(orig_zip, i);
        const char* name = zip_entry_name(orig_zip);

        // Skip copying the original file
        //if (std::string(name) != std::string("word/document.xml")) {
        if (std::string(name) != std::string("media/"))
        if (std::string(name) != std::string("content.xml")) {
            // Read the old content
            void* entry_buf;
            size_t entry_buf_size;
            zip_entry_read(orig_zip, &entry_buf, &entry_buf_size);

            // Write into new zip
            zip_entry_open(new_zip, name);
            zip_entry_write(new_zip, entry_buf, entry_buf_size);
            zip_entry_close(new_zip);

            free(entry_buf);
        }

        zip_entry_close(orig_zip);
    }

    // Close both zips
    zip_close(orig_zip);
    zip_close(new_zip);

    // Remove original zip, rename new to correct name
    //remove(original_file.c_str());
    rename(temp_file.c_str(), new_name.c_str());
}

duckx::Paragraph &duckx::Document::paragraphs() {
    //this->paragraph.set_parent(document.child("w:document").child("w:body"));
    this->paragraph.set_parent(document.child("office:document-content").child("office:body").child("office:text"));
    return this->paragraph;
}

duckx::Table &duckx::Document::tables() {
    //this->table.set_parent(document.child("w:document").child("w:body"));
    this->table.set_parent(document.child("office:document-content").child("office:body").child("office:text"));
    return this->table;
}

duckx::Style& duckx::Document::styles()
{
    this->style.set_parent(document.child("office:document-content").child("office:automatic-styles"));
    return this->style;
}

duckx::Table& duckx::Document::add_table(const std::string& stylename)
{
    pugi::xml_node new_table = this->document.child("office:document-content").child("office:body").append_child("table:table");
    new_table.append_attribute("table:style-name").set_value(stylename.c_str());

    return *new Table(new_table.parent(), new_table);

}

duckx::Paragraph& duckx::Document::add_paragraph(const std::string& stylename)
{
    pugi::xml_node new_paragraph = this->document.child("office:document-content").child("office:body").append_child("text:p");
    new_paragraph.append_attribute("text:style-name").set_value(stylename.c_str());

    return *new Paragraph(new_paragraph.parent(), new_paragraph);
}

duckx::Style::Style() {}

duckx::Style::Style(pugi::xml_node parent, pugi::xml_node current)
{
    this->set_parent(parent);
    this->set_current(current);
}

void duckx::Style::set_parent(pugi::xml_node node)
{
    this->parent = node;
    this->current = this->parent.child("style:style");
}

void duckx::Style::set_current(pugi::xml_node node) { this->current = node; }

bool duckx::Style::has_next() const { return this->current != 0; }

duckx::Style& duckx::Style::add_style(std::string stylename, duckx::styles st, std::vector<std::pair<std::string, std::string>> attr)
{
    // Add new run

    pugi::xml_node new_style = this->parent.append_child("style:style");
    new_style.append_attribute("style:name").set_value(stylename.c_str());
    pugi::xml_node new_style_props;
    switch (st)
    {
    case duckx::styles::table:
    {
        new_style.append_attribute("style:family").set_value("table");
        new_style_props = new_style.append_child("style:table-properties");
    }
        break;
    case duckx::styles::column:
    {
        new_style.append_attribute("style:family").set_value("table-column");
        new_style_props = new_style.append_child("style:table-column-properties");
    }
        break;
    case duckx::styles::row:
    {
        new_style.append_attribute("style:family").set_value("table-row");
        new_style_props = new_style.append_child("style:table-row-properties");
    }
        break;
    case duckx::styles::cell:
    {
        new_style.append_attribute("style:family").set_value("table-cell");
        new_style_props = new_style.append_child("style:table-cell-properties");
    }
        break;
    case duckx::styles::paragraph:
    {
        new_style.append_attribute("style:parent-style-name").set_value("RegPar");
        new_style.append_attribute("style:family").set_value("paragraph");
        new_style_props = new_style.append_child("style:paragraph-properties");
    }
        break;
    case duckx::styles::run:
    {
        new_style.append_attribute("style:family").set_value("text");
        new_style_props = new_style.append_child("style:text-properties");
        new_style.append_attribute("style:parent-style-name").set_value("RegParText");
    }
        break;
    }
    for (auto elem : attr)
    {
        new_style_props.append_attribute(elem.first.c_str()).set_value(elem.second.c_str());
    }

    return *new Style(this->parent, new_style);
}

duckx::Style& duckx::Style::next() {
    this->current = this->current.next_sibling("style:style");
    return *this;
}
