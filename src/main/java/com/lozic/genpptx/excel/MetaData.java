package com.lozic.genpptx.excel;

public class MetaData {
	// 字段名称
		private String field;
		// 字段描述
		private String comment;
		// 默认描述
		private String defaultComments;

		public String getDefaultComments() {
			return defaultComments;
		}

		public void setDefaultComments(String defaultComments) {
			this.defaultComments = defaultComments;
		}

		public MetaData(String field, String defaultComments) {
			this.field = field;
			this.defaultComments = defaultComments;
		}

		public String getField() {
			return field;
		}

		public void setField(String field) {
			this.field = field;
		}

		public String getComment() {
			return comment;
		}

		public void setComment(String comment) {
			this.comment = comment;
		}
		
		public String toString()
		{
			return field+"#"+comment+"#"+defaultComments;
		}
		
		public  boolean compareField(){
			if(defaultComments.equals(comment)){
			  return true;
			}
			return false;
		}

}
