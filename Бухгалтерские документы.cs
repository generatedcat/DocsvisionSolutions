using System;
using System.Text;
using System.Linq;

using System.Windows.Forms;
using System.Windows.Forms.Design;

using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.ObjectModel;
using DocsVision.Platform.ObjectModel.Search;

using DocsVision.BackOffice.WinForms;
using DocsVision.BackOffice.WinForms.Controls;
using DocsVision.BackOffice.WinForms.Design.LayoutItems;
using DocsVision.BackOffice.ObjectModel;
using DocsVision.BackOffice.ObjectModel.Services;
using DocsVision.BackOffice.ObjectModel.Services.Entities;
using DocsVision.BackOffice.ObjectModel.Services.Entities.KindSetting;
using DocsVision.BackOffice.WinForms.Design;
using DocsVision.Platform.Extensibility;
using Docsvision.DocumentsManagement;

using Microsoft.Office.Interop.Word;



namespace BackOffice
{
    public class CardDocumentБухгалтерские_документыScript : CardDocumentВнутренний_документScript
    {
    
    #region Properties
		
		protected ICustomizableControl customizable {get {return base.CardControl as ICustomizableControl;}}
		protected ObjectContext Context { get { return base.CardControl.ObjectContext; } }
		protected IUIService UIService { get { return Context.GetService<IUIService>(); } }
		protected ITaskService TaskService { get { return Context.GetService<ITaskService>(); } }
		protected ITaskListService TaskListService { get { return Context.GetService<ITaskListService>(); } }
		protected IStaffService StaffService { get { return this.Context.GetService<IStaffService>(); } }
		protected IStateService StateService { get { return this.Context.GetService<IStateService>(); } }
		protected IDocumentService DocumentService { get { return this.Context.GetService<IDocumentService>(); } }
		protected DocsVision.BackOffice.ObjectModel.Document Document { get { return base.CardControl.BaseObject as DocsVision.BackOffice.ObjectModel.Document; } }
		

    #endregion

    #region Methods
		
			
	private void setMyHint(){
		ILayoutPropertyItem item = customizable.FindPropertyItem<ILayoutPropertyItem>("Hint");
		if (item != null) {
			if (Document.SystemInfo.State.DefaultName=="Is registered"){
				item.ControlValue="Документ у сотрудника снабжения. Необходимо указать куратора и запустить документ на согласование";
			}
			else if (Document.SystemInfo.State.DefaultName=="CuratorApproving"){
				item.ControlValue="Документ у куратора. Куратору необходимо из задания на согласование указать следующего участника согласования";
			}
			else if (Document.SystemInfo.State.DefaultName=="TechDetailsApproval") item.ControlValue="Документ на согласовании у технического директора";
			else if (Document.SystemInfo.State.DefaultName=="FinancialApproval") item.ControlValue="Документ на согласовании у бухгалтера";
			else if (Document.SystemInfo.State.DefaultName=="DirectorApproving") item.ControlValue="Документ на согласовании у руководства";
			else if (Document.SystemInfo.State.DefaultName=="Performing") item.ControlValue="Документ на исполнении. Необходимо завершить все задания по документу, после чего документ может быть перемещен в архив";
			else if (Document.SystemInfo.State.DefaultName=="Returned") item.ControlValue="Документ возвращен сотруднику снабжения. Далее он может быть направлен повторно куратору,или на согласование технической или финансовой части";
			item.Commit();
		}
	}
	
	private void sendTasks(String Performer){
		try
        {
			//сохраняем карточку
			if (!SaveCard()){
				UIService.ShowError("Сохранение не удалось. Выполните сохранение вручную");
				return;
			}
            //this.Context.SaveObject(this.Document);
			Guid TaskKindID = new Guid("0F6C537B-9D75-43EC-9D91-B29C5F1B7C4D");   //айди типа задания на исполнение
			KindsCardKind kind = Context.GetObject<KindsCardKind>(TaskKindID);
			DocsVision.BackOffice.ObjectModel.Task oTask = this.TaskService.CreateTask(kind); 
			TaskService.InitializeDefaults(oTask);
			
			//заполняем данные задания
			oTask.MainInfo.Author = this.StaffService.GetCurrentEmployee();//текущий сотрудник
			//содержание задания
			oTask.MainInfo.Content = "Документ пришел Вам на исполнение/ознакомление";
			oTask.MainInfo.StartDate = DateTime.Now;
			oTask.MainInfo.OnControl = false;
			oTask.MainInfo.Name = "Задание по бухгалтерскому документу "+this.Document.MainInfo.Name;
			oTask.Description = "Задание по бухгалтерскому документу "+this.Document.MainInfo.Name;
			//определяем важность
			oTask.MainInfo.Priority = TaskPriority.Normal;
			
			BaseCardSectionRow contractRow = (BaseCardSectionRow)this.Document.GetSection(new System.Guid("3997861D-4FF5-496A-B8A2-D16617DE91D7"))[0];
			if (contractRow["ContractEnd"]!=null)
				oTask.MainInfo.EndDate = Convert.ToDateTime(contractRow["ContractEnd"]);
			
			
			//добавляем исполнителя
			StaffEmployee oPerf = Context.GetObject<StaffEmployee>(new System.Guid(Performer));;
			TaskService.AddSelectedPerformer(oTask.MainInfo, oPerf);
			
			TaskSetting oTaskSetting = TaskService.GetKindSettings(kind);
			//добавляем ссылку на родительскую карточку
			TaskService.AddLinkOnParentCard(oTask, oTaskSetting, this.Document);
			//добавляем ссылку на задание в карточку
			TaskListService.AddTask(this.Document.MainInfo.Tasks, oTask, this.Document);
			//создаем и сохраняем новый список заданий
			TaskList newTaskList = TaskListService.CreateTaskList();
			Context.SaveObject<DocsVision.BackOffice.ObjectModel.TaskList>(newTaskList);
			//записываем в задание
			oTask.MainInfo.ChildTaskList = newTaskList;
			Context.SaveObject<DocsVision.BackOffice.ObjectModel.Task>(oTask);
			Context.SaveObject<DocsVision.BackOffice.ObjectModel.Document>(this.Document);
			//MessageBox.Show("Документ готов, запускаем");
			//запускаем задание на исполнение
			string oErrMessageForStart;
			bool CanStart = TaskService.ValidateForBegin(oTask, out oErrMessageForStart);
			if (CanStart)
                {
					//MessageBox.Show("Can start");
                  TaskService.StartTask(oTask);
               
                  //MessageBox.Show("Изменяется состояние задания");
					StatesState oStartedState = StateService.GetStates(Context.GetObject<KindsCardKind>(TaskKindID)).FirstOrDefault(br => br.DefaultName == "Started");
                  oTask.SystemInfo.State = oStartedState;
                  
                  UIService.ShowMessage("Документ успешно отправлен пользователю "+oPerf.ToString(), "Отправка задания");
				  
                }
               
               
                else
                    UIService.ShowMessage(oErrMessageForStart, "Ошибка отправки задания");
         
            Context.SaveObject<DocsVision.BackOffice.ObjectModel.Task>(oTask);
   			   
            }
         
   			catch (System.Exception ex)
      		{
         		UIService.ShowError(ex);
      		}
	}
	
		
	
	private void addTaskRows(DocsVision.BackOffice.ObjectModel.Task currentTask){
		
		//добавить исполниетлей и ход согласования
		var DocumentCuratorSection = Document.GetSection(new System.Guid("281A97FF-667F-46C8-8FBE-7CFC02EDFEDB"));
		foreach (BaseCardSectionRow docRow in DocumentCuratorSection){
			var CuratorSection = currentTask.GetSection(new System.Guid("7BEB89E5-4A68-445A-85BA-EEEFC0118623")); // Секция Куратор Задания
			BaseCardSectionRow newRow = new BaseCardSectionRow();
			newRow["Curator"]=docRow["Approver"];
			CuratorSection.Add(newRow);
		}
		var TechDirectorSection = Document.GetSection(new System.Guid("F47D0D6B-07FE-4198-8F79-348AC55086E5"));
		foreach (BaseCardSectionRow docRow in TechDirectorSection){
			var CuratorSection = currentTask.GetSection(new System.Guid("947F528D-0400-428D-9403-3A6F76BFE4CE")); // Секция тех.директор
			BaseCardSectionRow newRow = new BaseCardSectionRow();
			newRow["Director"]=docRow["Confirm"];
			CuratorSection.Add(newRow);
		}
		var AccountantSection = Document.GetSection(new System.Guid("D9F3BB4C-9C1A-464C-90F3-3D9657864709")); // бухгалтер
		foreach (BaseCardSectionRow docRow in AccountantSection){
			var CuratorSection = currentTask.GetSection(new System.Guid("9D5136DB-59BD-4039-BC34-1D26227F0A34"));
			BaseCardSectionRow newRow = new BaseCardSectionRow();
			newRow["Accountant"]=docRow["Signer"];
			CuratorSection.Add(newRow);
		}
		var ManagerSection = Document.GetSection(new System.Guid("B6DFAEAD-BAAA-4024-908C-5DBD693D0FD3")); // руководитель
		foreach (BaseCardSectionRow docRow in ManagerSection){
			var CuratorSection = currentTask.GetSection(new System.Guid("52B4B182-0FEE-4CB4-886C-965C8CC71CDE"));
			BaseCardSectionRow newRow = new BaseCardSectionRow();
			newRow["Manager"]=docRow["ReceiverStaff"];
			CuratorSection.Add(newRow);
		}
		var PerformerSection = Document.GetSection(new System.Guid("AF798AE7-BAAC-486E-84EF-82C59DC00A7E")); // на ознакомление
		foreach (BaseCardSectionRow docRow in PerformerSection){
			var CuratorSection = currentTask.GetSection(new System.Guid("0E0E6967-2910-4D60-8132-B34F52DC1571")); //Секция тех.директор
			BaseCardSectionRow newRow = new BaseCardSectionRow();
			newRow["AcquaintancePersons"]=docRow["AcquaintancePersons"];
			CuratorSection.Add(newRow);
		}
		
		BaseCardSectionRow mainRow = (BaseCardSectionRow)this.Document.GetSection(new System.Guid("30EB9B87-822B-4753-9A50-A1825DCA1B74"))[0];//First row of document main section
		BaseCardSectionRow mainTaskRow = (BaseCardSectionRow)currentTask.GetSection(new System.Guid("20D21193-9F7F-4B62-8D69-272E78E1D6A8"))[0]; //First row of task main section
		mainTaskRow["Partially"]=mainRow["WasSent"];
		mainTaskRow["PartialPurchase"]=mainRow["Sum"];
		mainTaskRow["Comment"]=mainRow["ExternalNumber"];
		//информацию о ходе согласования из род документа
		foreach (BaseCardSectionRow row in this.Document.GetSection(new System.Guid("AACEF937-EAD2-4AFD-A64C-DC42D7846B80"))){ // журнал согласования
			var CuratorTaskSection = currentTask.GetSection(new System.Guid("4B787F44-FBBD-47C1-A883-D9518B7B06DB"));
			BaseCardSectionRow newRow = new BaseCardSectionRow();
			newRow["Approver"] = row["Employee"];
			newRow["Date"] = row["Date"];
			newRow["Result"] = row["Result"];
			newRow["Comment"] = row["Comment"];
			CuratorTaskSection.Add(newRow);
		}
		
		Context.SaveObject<DocsVision.BackOffice.ObjectModel.Task>(currentTask);
	}
		
	private void UpdateRegNumber()
	{
	    try
	    {
            ILayoutPropertyItem numberControl = customizable.FindPropertyItem<ILayoutPropertyItem>("RegNumber");
            if (numberControl == null){
				UIService.ShowMessage("Нет поля для заполнения регистрационного номера");
            	return;
			}
			if (!(numberControl.ControlValue is Guid)){
				UIService.ShowMessage("Поле регистрационный номер имеет неверный формат ввода данных");
            	return;
			}
			if (((Guid)numberControl.ControlValue) != Guid.Empty){
				return;
			}
			INumerationRulesService numerationService = CardControl.ObjectContext.GetService<INumerationRulesService>();
			NumerationRulesRule rule = CardControl.ObjectContext.FindObject<NumerationRulesRule>(new QueryObject("RuleName", "Бухгалтерские документы"));
            if (rule == null)
                  return;
			// собственно выдача номера и установка его в контроле
            BaseCardNumber number = numerationService.CreateNumber(this.CardData, this.BaseObject, rule);
            numberControl.ControlValue = CardControl.ObjectContext.GetObjectRef(number).Id;
            numberControl.Commit();
			
			this.Document.MainInfo.DeliveryDate = DateTime.Now.Date;
			
			Context.SaveObject<DocsVision.BackOffice.ObjectModel.Document>(Document);
		
        }
      catch (Exception ex)
      {
            UIService.ShowMessage(ex.Message);
      }

    }
	
	private void sendApproval(String role, StaffEmployee Performer, int performingType){
		KindsCardKind kind = Context.GetObject<KindsCardKind>(new Guid("F841AEE1-6018-44C6-A6A6-D3BAE4A3439F")); // Задание на согласование фин.документов
		DocsVision.BackOffice.ObjectModel.Task task = TaskService.CreateTask(kind);
		task.MainInfo.Name="Согласование бухгалтерских документов "+Document.MainInfo.Name;
		task.Description="Согласование бухгалтерских документов "+Document.MainInfo.Name;
		string content = "Вы назначены " + role + " при согласовании бухгалтерских документов.";
		content = content + " Пожалуйста, отметьте свое решение кнопками Согласован или Не согласован и напишите соответствующий комментарий";
		task.MainInfo.Content=content;
		task.MainInfo.Author = this.StaffService.GetCurrentEmployee();
		task.MainInfo.StartDate=DateTime.Now;
		task.MainInfo.Priority=TaskPriority.Normal;
		task.Preset.AllowDelegateToAnyEmployee=true;
		TaskService.AddSelectedPerformer(task.MainInfo, Performer);
		BaseCardSectionRow taskRow = (BaseCardSectionRow)task.GetSection(new System.Guid("20D21193-9F7F-4B62-8D69-272E78E1D6A8"))[0];
		taskRow["PerformanceType"]=performingType;
		addTaskRows(task);
		//добавляем ссылку на родительскую карточку
		TaskService.AddLinkOnParentCard(task, TaskService.GetKindSettings(kind), this.Document);
		//добавляем ссылку на задание в карточку
		TaskListService.AddTask(this.Document.MainInfo.Tasks, task, this.Document);
		//создаем и сохраняем новый список заданий
		TaskList newTaskList = TaskListService.CreateTaskList();
		Context.SaveObject<DocsVision.BackOffice.ObjectModel.TaskList>(newTaskList);
		//записываем в задание
		task.MainInfo.ChildTaskList = newTaskList;
		Context.SaveObject<DocsVision.BackOffice.ObjectModel.Task>(task);
		string oErrMessageForStart;
		bool CanStart = TaskService.ValidateForBegin(task, out oErrMessageForStart);
		if (CanStart){
			TaskService.StartTask(task);
            StatesState oStartedState = StateService.GetStates(kind).FirstOrDefault(br => br.DefaultName == "Started");
            task.SystemInfo.State = oStartedState;
              
            UIService.ShowMessage("Документ успешно отправлен пользователю "+Performer.DisplayString, "Отправка задания");
			  
        }
		
        else
            UIService.ShowMessage(oErrMessageForStart, "Ошибка отправки задания");
			
		Context.SaveObject<DocsVision.BackOffice.ObjectModel.Task>(task);
	}

    #endregion

    #region Event Handlers
	
	override public void CardActivated(DocsVision.Platform.WinForms.CardActivatedEventArgs e)
	{
		DocumentHelper.CardActivated(e);

		setMyHint();
	}
	
	override public void CardSaved()
	{
		UpdateRegNumber();
		DocumentHelper.CardSaved();	
	}
	
	private void командаНаДоработку_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e){
		Context.SaveObject<DocsVision.BackOffice.ObjectModel.Document>(this.Document);
		//MessageBox.Show("in work buttom clicked");
		SectionData acquaintanceStaffSection = this.CardData.Sections[this.CardData.Type.Sections["AcquaintanceStaff"].Id];
		foreach(RowData row in acquaintanceStaffSection.Rows)
		{
			//MessageBox.Show(row["AcquaintancePersons"].ToString());
			sendTasks(row["AcquaintancePersons"].ToString());
		}
	}
    private void командаВРаботу_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
    {
        if (!SaveCard()){
			UIService.ShowError("Сохранение не удалось. Выполните сохранение вручную");
			return;
		}
		if (Document.SystemInfo.State.DefaultName=="Is registered"){
			var curatorRow = Document.GetSection(new System.Guid("281A97FF-667F-46C8-8FBE-7CFC02EDFEDB")); //Секция Согласующие - соответствует Куратору
			if (curatorRow.Count==0){
				UIService.ShowMessage("Не заполнено поле Куратор", "Проверка исполнителей задания на согласование");
				return;
			}
			int performingType = 0;
			if (curatorRow.Count>1)performingType = 1;
			foreach(BaseCardSectionRow row in curatorRow){
				sendApproval("куратором", Context.GetObject<StaffEmployee>(new System.Guid(row["Approver"].ToString())), performingType);
			}
			changeState("CuratorApproving");
		}
		else if (Document.SystemInfo.State.DefaultName=="FinancialApproval"){
			var managerRow = Document.GetSection(new System.Guid("B6DFAEAD-BAAA-4024-908C-5DBD693D0FD3")); //Секция Получатель - соответствует Руководителю
			if (managerRow.Count==0){
				UIService.ShowMessage("Не заполнено поле руководитель", "Проверка исполнителей задания на согласование");
				return;
			}
			int performingType = 0;
			if (managerRow.Count>1)performingType = 1;
			foreach(BaseCardSectionRow row in managerRow){
				sendApproval("руководителем", Context.GetObject<StaffEmployee>(new System.Guid(row["ReceiverStaff"].ToString())), performingType);
			}
			changeState("DirectorApproving");
		}
		else {
			UIService.ShowMessage("Состояние документа не предусматривает дальнейшую отправку на согласование по данной кнопке", "Проверка состояния документа");
		}
		
		
    }
	
	private void командаБухгалтеру_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
    {
		if (!SaveCard()){
			UIService.ShowError("Сохранение не удалось. Выполните сохранение вручную");
			return;
		}
		
		var accountantRow = Document.GetSection(new System.Guid("D9F3BB4C-9C1A-464C-90F3-3D9657864709")); //Секция подписант - соответствует бухгалтеру
		if (accountantRow.Count==0){
			UIService.ShowMessage("Не заполнено поле бухгалтер", "Проверка исполнителей задания на согласование");
			return;
		}
		BaseCardSectionRow row = (BaseCardSectionRow)accountantRow[0];
		sendApproval("бухгалтером", Context.GetObject<StaffEmployee>(new System.Guid(row["Signer"].ToString())), 0);
		changeState("FinancialApproval");
	}
	
	private void printApproval_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
	{
		DocsVision.BackOffice.ObjectModel.Document FileCardDocument = this.Context.GetObject<DocsVision.BackOffice.ObjectModel.Document>(new System.Guid("224D391A-B7ED-E511-9414-E4115BB03AEA"));
		//MessageBox.Show("Документ нашли");
		string oTempPath = System.IO.Path.GetTempPath(); 
		string oFilePath = oTempPath+FileCardDocument.MainInfo.MainFileName;
		this.DocumentService.DownloadMainFile(FileCardDocument, oFilePath);
							
		//MessageBox.Show("Документ скачали");
		//получаем объект word
		Microsoft.Office.Interop.Word.ApplicationClass oApplication = new Microsoft.Office.Interop.Word.ApplicationClass();
		oApplication.Visible=true;
		object oMissing = Type.Missing;
		object oFileNameMain = (object)oFilePath;
		object oFalse = false;
		object oTrue = true;
				
		try{
		Microsoft.Office.Interop.Word.Document oMainFileWord = oApplication.Documents.Open(ref oFileNameMain, ref oMissing, 
					ref oFalse,  ref oFalse, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, 
					ref oMissing, ref oTrue, ref oMissing, ref oMissing, ref oTrue, ref oMissing);
			
			//MessageBox.Show("Открыли");
	
			object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
		
			string temp =String.Empty;
		
			BaseCardSectionRow row = (BaseCardSectionRow)Document.GetSection(new System.Guid("30EB9B87-822B-4753-9A50-A1825DCA1B74"))[0];
			Guid tempId = new System.Guid(row["ResponsDepartment"].ToString());
			StaffUnit unit = Context.GetObject<StaffUnit>(tempId);
				
			oApplication.Selection.Find.ClearFormatting();
			oApplication.Selection.Find.Text = "&Организация.";
			oApplication.Selection.Find.Replacement.ClearFormatting();
			oApplication.Selection.Find.Replacement.Text = unit.Name;
			oApplication.Selection.Find.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref replaceAll, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
			
			oApplication.Selection.Find.ClearFormatting();
			oApplication.Selection.Find.Text = "&Состояние.";
			oApplication.Selection.Find.Replacement.ClearFormatting();
			oApplication.Selection.Find.Replacement.Text = Document.SystemInfo.State.LocalizedName;
			oApplication.Selection.Find.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref replaceAll, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
			
			BaseCardSectionRow docPartnerRow = (BaseCardSectionRow)Document.GetSection(new System.Guid("6E976D72-3EA7-4708-A2C2-2A1499141301"))[0];
			CardData PartnerDictionary=Session.CardManager.GetCardData(new System.Guid("65FF9382-17DC-4E9F-8E93-84D6D3D8FE8C")); //Справочник контрагентов
			RowData PartnerRow =PartnerDictionary.Sections[new System.Guid("C78ABDED-DB1C-4217-AE0D-51A400546923")].GetRow(new System.Guid(docPartnerRow["SenderOrg"].ToString()));
			oApplication.Selection.Find.ClearFormatting();
			oApplication.Selection.Find.Text = "&Контрагент.";
			oApplication.Selection.Find.Replacement.ClearFormatting();
			oApplication.Selection.Find.Replacement.Text = PartnerRow.GetString("Name");
			oApplication.Selection.Find.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref replaceAll, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

			if (row["StatusId"]!=null){
				CardData dictionaryCard = this.Session.CardManager.GetCardData(new System.Guid("4538149D-1FC7-4D41-A104-890342C6B4F8"));
				SectionData dictionarySection = dictionaryCard.Sections[new System.Guid("A1DCE6C1-DB96-4666-B418-5A075CDB02C9")];
				RowData rdItemTypesNuno = dictionarySection.GetRow(new System.Guid("10630364-D63F-4249-83E8-5DE95B2CD385")); //виды бухгалтерских документов
				foreach (RowData rdItems in rdItemTypesNuno.ChildSections[new System.Guid("1B1A44FB-1FB1-4876-83AA-95AD38907E24")].Rows){
					if (row["StatusId"].ToString().ToUpper()=="{"+rdItems.Id.ToString().ToUpper()+"}"){
						temp = rdItems.GetString("Name");
					}
				}
			}
			else temp = String.Empty;
			oApplication.Selection.Find.ClearFormatting();
			oApplication.Selection.Find.Text = "&Вид документа.";
			oApplication.Selection.Find.Replacement.ClearFormatting();
			oApplication.Selection.Find.Replacement.Text = temp;
			oApplication.Selection.Find.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref replaceAll, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
			
			BaseCardSectionRow contractRow = (BaseCardSectionRow)Document.GetSection(new System.Guid("3997861D-4FF5-496A-B8A2-D16617DE91D7"))[0];
			temp = contractRow["AddAgreementNumber"]+"";
			
			oApplication.Selection.Find.ClearFormatting();
			oApplication.Selection.Find.Text = "&Номер.";
			oApplication.Selection.Find.Replacement.ClearFormatting();
			oApplication.Selection.Find.Replacement.Text = temp;
			oApplication.Selection.Find.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref replaceAll, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
			
			oApplication.Selection.Find.ClearFormatting();
			oApplication.Selection.Find.Text = "&Дата.";
			oApplication.Selection.Find.Replacement.ClearFormatting();
			oApplication.Selection.Find.Replacement.Text = Convert.ToDateTime(contractRow["AttachmentDate"]+"").ToShortDateString();
			oApplication.Selection.Find.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref replaceAll, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
			
			oApplication.Selection.Find.ClearFormatting();
			oApplication.Selection.Find.Text = "&Сумма.";
			oApplication.Selection.Find.Replacement.ClearFormatting();
			oApplication.Selection.Find.Replacement.Text = contractRow["MySum"]+"";
			oApplication.Selection.Find.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref replaceAll, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
			
			string partially;
			if (row["WasSent"]+""=="False"){
				temp = contractRow["MySum"]+"";
				partially = "Нет";
			}
			else{
				temp = row["Sum"]+"";
				partially = "Да";
			}
			
			oApplication.Selection.Find.ClearFormatting();
			oApplication.Selection.Find.Text = "&К оплате.";
			oApplication.Selection.Find.Replacement.ClearFormatting();
			oApplication.Selection.Find.Replacement.Text = temp;
			oApplication.Selection.Find.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref replaceAll, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
			
			if (row["SLA"]!=null){
				//MessageBox.Show("якобы условия оплаты заполнены"+row["SLA"]);
				CardData dictionaryCard = this.Session.CardManager.GetCardData(new System.Guid("4538149D-1FC7-4D41-A104-890342C6B4F8"));
				SectionData dictionarySection = dictionaryCard.Sections[new System.Guid("A1DCE6C1-DB96-4666-B418-5A075CDB02C9")];
				RowData rdItemTypesNuno = dictionarySection.GetRow(new System.Guid("305EF160-0CBA-41B7-8187-E6383A471849")); //условия оплаты
				foreach (RowData rdItems in rdItemTypesNuno.ChildSections[new System.Guid("1B1A44FB-1FB1-4876-83AA-95AD38907E24")].Rows){
					if (row["StatusId"].ToString().ToUpper()=="{"+rdItems.Id.ToString().ToUpper()+"}"){
						temp = rdItems.GetString("Name");
					}
				}
			}
			else{
				//MessageBox.Show("условия оплаты не заполнены");
				temp = String.Empty;
			}
			oApplication.Selection.Find.ClearFormatting();
			oApplication.Selection.Find.Text = "&Условия.";
			oApplication.Selection.Find.Replacement.ClearFormatting();
			oApplication.Selection.Find.Replacement.Text = temp;
			oApplication.Selection.Find.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref replaceAll, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
			
			oApplication.Selection.Find.ClearFormatting();
			oApplication.Selection.Find.Text = "&Валюта.";
			oApplication.Selection.Find.Replacement.ClearFormatting();
			oApplication.Selection.Find.Replacement.Text = contractRow["ContractCurrency"]+"";
			oApplication.Selection.Find.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref replaceAll, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
			
			oApplication.Selection.Find.ClearFormatting();
			oApplication.Selection.Find.Text = "&Частично.";
			oApplication.Selection.Find.Replacement.ClearFormatting();
			oApplication.Selection.Find.Replacement.Text = partially;
			oApplication.Selection.Find.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
	                ref replaceAll, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
			
			Microsoft.Office.Interop.Word.Table newTable;
			Microsoft.Office.Interop.Word.Range myRange = oMainFileWord.Range();
			//object startOfRange = myRange.Text.IndexOf("&Таблица.");
			//object endOfRange = myRange.Text.IndexOf("&Таблица.") + 9;
			//Microsoft.Office.Interop.Word.Range wrdRng = oMainFileWord.Range(ref startOfRange, ref endOfRange);
			object oEndOfDoc = "\\endofdoc";
			Microsoft.Office.Interop.Word.Range wrdRng =  oMainFileWord.Bookmarks.get_Item(ref oEndOfDoc).Range;
			//MessageBox.Show("range is determined");
	        newTable = oMainFileWord.Tables.Add(wrdRng, 1, 4, ref oMissing, ref oMissing);
			//MessageBox.Show("table added");
	        newTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
	        newTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
	        newTable.AllowAutoFit = true;
			newTable.Cell(1,1).Range.Text = "Сотрудник";
			newTable.Cell(1,1).Range.Font.Bold=1;
			newTable.Cell(1,2).Range.Text = "Дата согласования";
			newTable.Cell(1,2).Range.Bold = 1;
			newTable.Cell(1,3).Range.Text = "Результат";
			newTable.Cell(1,3).Range.Bold = 1;
			newTable.Cell(1,4).Range.Text = "Комментарий";
			newTable.Cell(1,4).Range.Bold = 1;
		
		int i=1;
		while (i<=Document.GetSection(new System.Guid("AACEF937-EAD2-4AFD-A64C-DC42D7846B80")).Count){
			newTable.Rows.Add();
			BaseCardSectionRow commentTable = (BaseCardSectionRow)Document.GetSection(new System.Guid("AACEF937-EAD2-4AFD-A64C-DC42D7846B80"))[i-1];
			string employeeId = commentTable["Employee"].ToString();
			StaffEmployee employee = Context.GetObject<StaffEmployee>(new System.Guid(employeeId));
			newTable.Cell(i+1,1).Range.Text = employee.DisplayString;
			newTable.Cell(i+1,1).Range.Bold = 0;
			newTable.Cell(i+1,2).Range.Text = Convert.ToDateTime(commentTable["Date"]+"").ToShortDateString();
			if (commentTable["Result"]+""=="0")
				temp = "Согласовано";
			else
				temp="Не согласовано";
			newTable.Cell(i+1,2).Range.Bold = 0;
			newTable.Cell(i+1,3).Range.Text = temp;
			newTable.Cell(i+1,3).Range.Bold = 0;
			newTable.Cell(i+1,4).Range.Text = commentTable["Comment"]+"";
			newTable.Cell(i+1,4).Range.Bold = 0;
			i++;
		}
		}
		catch (Exception ex){
			UIService.ShowError(ex, "Во время создания файла произошла ошибка");
		}
				
		/*удал¤ем временный файл
		if (System.IO.File.Exists(oFilePath))
        System.IO.File.Delete(oFilePath);*/

		//закрываем ворд
		//oApplication.Quit(ref oMissing, ref oMissing, ref oMissing);
	}

	private void ToTechDirector_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e){
		if (!SaveCard()){
			UIService.ShowError("Сохранение не удалось. Выполните сохранение вручную");
			return;
		}
		
		var techDirectorRow = Document.GetSection(new System.Guid("F47D0D6B-07FE-4198-8F79-348AC55086E5")); //Секция утверждено - соответствует тех.диру
		if (techDirectorRow.Count==0){
			UIService.ShowMessage("Не заполнено поле технического директора", "Проверка исполнителей задания на согласование");
			return;
		}
		int performingType = 0;
		foreach(BaseCardSectionRow row in techDirectorRow){
			sendApproval("техническим директором", Context.GetObject<StaffEmployee>(new System.Guid(row["Confirm"].ToString())), performingType);
		}
	}
	private void командаТехДиректору_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
    {
		if (!SaveCard()){
			UIService.ShowError("Сохранение не удалось. Выполните сохранение вручную");
			return;
		}
		
		var techDirectorRow = Document.GetSection(new System.Guid("F47D0D6B-07FE-4198-8F79-348AC55086E5")); //Секция утверждено - соответствует тех.диру
		if (techDirectorRow.Count==0){
			UIService.ShowMessage("Не заполнено поле технического директора", "Проверка исполнителей задания на согласование");
			return;
		}
		int performingType = 0;
		if (techDirectorRow.Count>1)performingType = 1;
		foreach(BaseCardSectionRow row in techDirectorRow){
			sendApproval("техническим директором", Context.GetObject<StaffEmployee>(new System.Guid(row["Confirm"].ToString())), performingType);
		}
		changeState("TechDetailsApproval");
	}
    private void командаЗарегистрировать_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
    {
		if (!SaveCard()){
			UIService.ShowError("Сохранение не удалось. Выполните сохранение вручную");
			return;
		}
        KindsCardKind kind = Context.GetObject<KindsCardKind>(new Guid("AB801854-70AF-4B6C-AB48-1B59B5D11AA9"));
		DocsVision.BackOffice.ObjectModel.Task task = TaskService.CreateTask(kind);
		task.MainInfo.Name="Проверка бухгалтерских документов "+Document.MainInfo.Name;
		task.Description="Проверка бухгалтерских документов "+Document.MainInfo.Name;
		string content = "Вы назначены ответственным за первичную проверку бухгалтерских документов и выставление согласущего куратора.";
		content = content + ". Пожалуйста, отметьте в родительском документе сотрудника в поле Куратор и отправьте документ на согласование";
		task.MainInfo.Content=content;
		task.MainInfo.Author = this.StaffService.GetCurrentEmployee();
		task.MainInfo.StartDate=DateTime.Now;
		task.MainInfo.Priority=TaskPriority.Normal;
		task.Preset.AllowDelegateToAnyEmployee=true;
		TaskService.AddSelectedPerformer(task.MainInfo, Document.MainInfo.Registrar);
		//добавляем ссылку на родительскую карточку
		TaskService.AddLinkOnParentCard(task, TaskService.GetKindSettings(kind), this.Document);
		//добавляем ссылку на задание в карточку
		TaskListService.AddTask(this.Document.MainInfo.Tasks, task, this.Document);
		//создаем и сохраняем новый список заданий
		TaskList newTaskList = TaskListService.CreateTaskList();
		Context.SaveObject<DocsVision.BackOffice.ObjectModel.TaskList>(newTaskList);
		//записываем в задание
		task.MainInfo.ChildTaskList = newTaskList;
		Context.SaveObject<DocsVision.BackOffice.ObjectModel.Task>(task);
		string oErrMessageForStart;
		bool CanStart = TaskService.ValidateForBegin(task, out oErrMessageForStart);
		if (CanStart){
			TaskService.StartTask(task);
            StatesState oStartedState = StateService.GetStates(kind).FirstOrDefault(br => br.DefaultName == "Started");
            task.SystemInfo.State = oStartedState;
              
            UIService.ShowMessage("Документ успешно отправлен пользователю "+Document.MainInfo.Registrar.DisplayString, "Отправка задания");
			  
        }
		
        else
            UIService.ShowMessage(oErrMessageForStart, "Ошибка отправки задания");
			
		Context.SaveObject<DocsVision.BackOffice.ObjectModel.Task>(task);
		
		changeState("Is registered");
    }

    private void Автор_ControlValueChanged(System.Object sender, System.EventArgs e)
    {
        BaseCardSectionRow row = (BaseCardSectionRow)Document.GetSection(new System.Guid("30EB9B87-822B-4753-9A50-A1825DCA1B74"))[0];
		if (this.Document.MainInfo.Author!=null){
			StaffUnit unit = getOrganisation(this.Document.MainInfo.Author.Unit);
			if (unit!=null){
				row["ResponsDepartment"]= CardControl.ObjectContext.GetObjectRef(unit).Id;
			}
		}
    }

    #endregion

    }
}
