using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Text;

namespace LAT.WorkflowUtilities.Note
{
	public sealed class DeleteAttachment : CodeActivity
	{
		[RequiredArgument]
		[Input("Note With Attachments To Remove")]
		[ReferenceTarget("annotation")]
		public InArgument<EntityReference> NoteWithAttachment { get; set; }

		[Input("Delete >= Than X Bytes (Empty = 0)")]
		public InArgument<int> DeleteSizeMax { get; set; }

		[Input("Delete <= Than X Bytes (Empty = 2,147,483,647)")]
		public InArgument<int> DeleteSizeMin { get; set; }

		[Input("Limit To Extensions (Comma Delimited, Empty = Ignore)")]
		public InArgument<string> Extensions { get; set; }

		[RequiredArgument]
		[Input("Add Delete Notice As Text?")]
		[Default("false")]
		public InArgument<bool> AppendNotice { get; set; }

		[Output("Number Of Attachments Deleted")]
		public OutArgument<int> NumberOfAttachmentsDeleted { get; set; }

		protected override void Execute(CodeActivityContext executionContext)
		{
			ITracingService tracer = executionContext.GetExtension<ITracingService>();
			IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
			IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
			IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

			try
			{
				EntityReference noteWithAttachment = NoteWithAttachment.Get(executionContext);
				int deleteSizeMax = DeleteSizeMax.Get(executionContext);
				int deleteSizeMin = DeleteSizeMin.Get(executionContext);
				string extensions = Extensions.Get(executionContext);
				bool appendNotice = AppendNotice.Get(executionContext);

				if (deleteSizeMax == 0) deleteSizeMax = int.MaxValue;
				if (deleteSizeMin > deleteSizeMax)
				{
					tracer.Trace("Exception: {0}", "Min:" + deleteSizeMin + " Max:" + deleteSizeMax);
					throw new InvalidPluginExecutionException("Minimum Size Cannot Be Greater Than Maximum Size");
				}

				Entity note = GetNote(service, noteWithAttachment.Id);
				if (!CheckForAttachment(note))
					return;

				string[] filetypes = new string[0];
				if (!string.IsNullOrEmpty(extensions))
					filetypes = extensions.Replace(".", string.Empty).Split(',');

				StringBuilder notice = new StringBuilder();
				int numberOfAttachmentsDeleted = 0;

				bool delete = false;

				if (note.GetAttributeValue<int>("filesize") >= deleteSizeMax)
					delete = true;

				if (note.GetAttributeValue<int>("filesize") <= deleteSizeMin)
					delete = true;

				if (filetypes.Length > 0 && delete)
					delete = ExtensionMatch(filetypes, note.GetAttributeValue<string>("filename"));

				if (delete)
				{
					numberOfAttachmentsDeleted++;

					if (appendNotice)
						notice.AppendLine("Deleted Attachment: " + note.GetAttributeValue<string>("filename") + " " +
										  DateTime.Now.ToShortDateString());

					UpdateNote(service, note, notice.ToString());
				}

				NumberOfAttachmentsDeleted.Set(executionContext, numberOfAttachmentsDeleted);
			}
			catch (Exception ex)
			{
				tracer.Trace("Exception: {0}", ex.ToString());
			}
		}

		private static bool CheckForAttachment(Entity note)
		{
			object oIsAttachment;
			bool hasValue = note.Attributes.TryGetValue("isdocument", out oIsAttachment);
			if (!hasValue)
				return false;

			return (bool)oIsAttachment;
		}

		private static Entity GetNote(IOrganizationService service, Guid noteId)
		{
			return service.Retrieve("annotation", noteId, new ColumnSet("filename", "filesize", "isdocument", "notetext"));
		}

		private void UpdateNote(IOrganizationService service, Entity note, string notice)
		{
			Entity updateNote = new Entity("annotation");
			updateNote.Id = note.Id;
			if (!string.IsNullOrEmpty(notice))
			{
				string newText = note.GetAttributeValue<string>("notetext");
				if (!string.IsNullOrEmpty(newText))
					newText += "\r\n";

				updateNote["notetext"] = newText + notice;
			}
			updateNote["isdocument"] = false;
			updateNote["filename"] = null;
			updateNote["documentbody"] = null;
			updateNote["filesize"] = null;

			service.Update(updateNote);
		}

		private static bool ExtensionMatch(IEnumerable<string> extenstons, string filename)
		{
			foreach (string ex in extenstons)
			{
				if (filename.EndsWith("." + ex))
					return true;
			}
			return false;
		}
	}
}