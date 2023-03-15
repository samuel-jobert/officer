#' @export
#' @title Removes a bookmark and its contents.
#' If start and end in same paragraph, removes paragraph
#' Else, removes all paragraphs between bookmark start and end (inluding start and end paragraphs)
#' @param x The officer word document
#' @param bkm The bookmark name to remove
#' @examples
#' mydoc <- read_docx("template_for_bookmark_remove.docx")
#' bkm = "bkm_in_paragraph"
#' # bkm = "box_bkm_in_paragraph"
#' # bkm = "bkm_multi_line"
#' # bkm = "box_bkm_multi_line"
#' mydoc <- mydoc %>% remove_bookmark(bkm)
#' print(mydoc, target = "bookmark_removed.docx")
remove_bookmark <- function(x, bkm) {
	
	if (!inherits(x, "rdocx")) {
		stop("Not a rdocx object.")
	}
	
	# Get bkm start node
	expr <- sprintf("/descendant::w:bookmarkStart[@w:name='%s']", bkm)
	bkm_start_node <- xml_find_first(x$doc_obj$get(), expr)
	# If bookmark not found, stop here
	if (inherits(bkm_start_node, "xml_missing")) {
		print("Bookmark ", bkm, " not found !")
		return(x)
	}
	
	# Bookmark end node
	bkm_id <- bkm_start_node %>% xml_attr("id")
	expr <- sprintf("/descendant::w:bookmarkEnd[@w:id='%s']", bkm_id)
	bkm_end_node <- xml_find_first(x$doc_obj$get(), expr)
	
	# Find parent of bookmark start & end (paragraph)
	bkm_start_parent <- bkm_start_node %>% xml_parent()
	bkm_end_parent <- bkm_end_node %>% xml_parent()

	# Get parent paraIds
	bkm_start_parent_id <- bkm_start_parent %>% xml_attr("paraId")
	bkm_end_parent_id <- bkm_end_parent %>% xml_attr("paraId")

	# Check if parents of parent paragraph are body or not (if body, parent paragraph is in x$officer_cursor)
	# If they are not body, they may be w:txbxContent (in text box)
	bkm_start_parent_parent <- bkm_start_parent %>% xml_parent() %>% xml_name()
	bkm_end_parent_parent <- bkm_end_parent %>% xml_parent() %>% xml_name()
	
	# Start and end in same par ?
	same_par <- (bkm_start_parent_id == bkm_end_parent_id)

	# If start and end parents have the same id (cursor in same paragraph) and 
	# parent paragraphs are 'body', we can use the body_remove function
	if (same_par & bkm_start_parent_parent == "body") {
		x <- x %>% cursor_bookmark(bkm)
		x <- x %>% body_remove()
		return(x)
	}
	# If same paragraph but parent not in officer_cursor
	# only remove parent paragraph
	if (same_par & bkm_start_parent_parent != "body") {
		xml_remove(bkm_start_parent)
		return(x)
	}
	# If not same paragraph, but paragraphs are in officer_cursor
	# delete officer cursors & remove nodes from start parent to end parent
	if (!same_par & bkm_start_parent_parent == "body") {
		expr <- "/w:document/w:body/*"
		doc_nodes <- xml_find_all(x$doc_obj$get(), expr)
		whichs <- which(doc_nodes %>% xml_attr("paraId") %in% c(bkm_start_parent_id, bkm_end_parent_id))
		x$officer_cursor$which <- whichs[1]
		for (cur_which in seq.int(whichs[1], whichs[2])){
			x$officer_cursor$nodes_names <- x$officer_cursor$nodes_names[-x$officer_cursor$which]
			xml_remove(doc_nodes[cur_which])
		}
		if (x$officer_cursor$which > length(x$officer_cursor$nodes_names)) {
			x$officer_cursor$which <- x$officer_cursor$which -1L
		}
		return(x)
	}
	# If not same paragraph and not in parents not officer cursors
	# remove nodes from start parent to end parent
	if (!same_par & bkm_start_parent_parent != "body") {
		nodes_id_to_remove <- bkm_start_parent_id
		next_ <- bkm_start_parent
		ok <- FALSE
		while (!ok) {
			# Fetch next parent sibling
			expr <- sprintf("//w:p[@w14:paraId='%s']//following-sibling::w:p", nodes_id_to_remove[length(nodes_id_to_remove)])
			next_ <- xml_find_first(next_, expr)
			nodes_id_to_remove <- append(nodes_id_to_remove, next_ %>% xml_attr("paraId"))
			# If next parent sibling is parent of bookmarkEnd, ok stop
			if (nodes_id_to_remove[length(nodes_id_to_remove)] == bkm_end_parent_id) {
				ok <- TRUE
			}
		}
		# Find all paragraph nodes with these paraIds
		expr <- paste0("//w:p[@w14:paraId='", nodes_id_to_remove, "']", collapse = "|")
		nodes_to_remove <- xml_find_all(x$doc_obj$get(), expr)
		xml_remove(nodes_to_remove)
		return(x)
	}
}


#' @export
#' @title Fix duplicated bookmark ids in officer docx document.
#' Sometimes, Word creates duplicated bookmark ids.
#' This brings problems when inserting docx file at cursor_bookmark() or 
#' using functions *_at_bkm
#' This function checks for duplicates and gives them disctinct new ones
#' @param x The officer word document
#' @examples
#' mydoc <- read_docx("doc_with_duplicated_bkm_ids.docx")
#' mydoc <- mydoc %>% fix_bookmark_ids()
#' # Do what you want with mydoc
fix_bookmark_ids <- function(x) {

	if (!inherits(x, "rdocx")) {
		stop("Not a rdocx object.")
	}
	
	expr <- sprintf("/descendant::w:bookmarkStart|/descendant::w:bookmarkEnd")
	nodes_with_bkm <- xml_find_all(x$doc_obj$get(), expr)
	expr <- sprintf("/descendant::w:bookmarkStart")
	nodes_with_bkm_start <- xml_find_all(nodes_with_bkm, expr)
	expr <- sprintf("/descendant::w:bookmarkEnd")
	nodes_with_bkm_end <- xml_find_all(nodes_with_bkm, expr)
	
	# If no bookmark found, do nothing and stop here
	if (inherits(nodes_with_bkm_end, "xml_missing")) {
		return(x)
	}

	# Fetch bkm ids
	bkm_ids <- sapply(nodes_with_bkm_start, function(node) {
		node %>% xml_attr("id")
	})
	duplicated_ids <- bkm_ids[duplicated(bkm_ids)]
	
	# No duplicated : do nothing and stop here
	if (length(duplicated_ids) == 0) {
		return(x)
	}
	max_id <- max(as.integer(bkm_ids))

	for (id in duplicated_ids) {
		max_id <- max_id + 1
		expr <- sprintf("//w:bookmarkStart[@w:id='%s']", id)
		bkm_start_to_modify <- xml_find_all(nodes_with_bkm_start, expr)[[2]]
		expr <- sprintf("//w:bookmarkEnd[@w:id='%s']", id)
		bkm_end_to_modify <- xml_find_all(nodes_with_bkm_end, expr)[[2]]
		xml_set_attr(bkm_start_to_modify, "w:id", as.character(max_id))
		xml_set_attr(bkm_end_to_modify, "w:id", as.character(max_id))
	}
	return(x)
}
