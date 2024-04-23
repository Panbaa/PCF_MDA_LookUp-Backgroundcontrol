import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class InputFieldLookUpToChange implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _container: HTMLDivElement;
	private _context: ComponentFramework.Context<IInputs>;
	private _notifyOutputChanged: () => void;
	private _inputData: HTMLLabelElement;
	private _contentContainer: HTMLDivElement;
	private _lookUpItem: HTMLDivElement;
	private _entityType = "";
	private _defaultViewId = "";
	private _selectedItem: ComponentFramework.LookupValue | undefined;
	private _colorInRgb: string;

	constructor() { }

	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
		this._container = container;
		this._notifyOutputChanged = notifyOutputChanged;

		// Cache information necessary for lookupObjects API based on control context
		this._entityType = context.parameters.controlValue.getTargetEntityType();
		this._defaultViewId = context.parameters.controlValue.getViewId();

		this._contentContainer = document.createElement("div");
		this._contentContainer.setAttribute("id", "lookUp");

		this._lookUpItem = document.createElement("div");
		this._lookUpItem.setAttribute("id", "lookUpItem");

		// Add element to display raw entity data
		this._inputData = document.createElement("label");
		this._inputData.setAttribute("id", "lookUpLabel");
		this._inputData.textContent = "";
		this._lookUpItem.appendChild(this._inputData);

		// Add element to remove entity data
		const deleteButton = document.createElement("button");
		deleteButton.setAttribute("id", "lookUpDeleteButton");
		deleteButton.onclick = () => this.clearLookupField();
		this._lookUpItem.appendChild(deleteButton);

		this._contentContainer.appendChild(this._lookUpItem);

		// Add element to display search-icon
		const lookUpIcon = document.createElement("div");
		lookUpIcon.setAttribute("id", "lookUpIcon");
		this._contentContainer.appendChild(lookUpIcon);

		// Add button to trigger lookupObjects API for the primary lookup property
		const lookupObjectsButton = document.createElement("button");
		lookupObjectsButton.setAttribute("id", "lookUpButton");
		lookupObjectsButton.onclick = this.performLookupObjects.bind(this, this._entityType, this._defaultViewId, (value, update = true) => {
			this._selectedItem = value;
		});
		this._contentContainer.appendChild(lookupObjectsButton);

		this._container.appendChild(this._contentContainer);

		this._contentContainer.style.backgroundColor = this._colorInRgb;
		this._lookUpItem.style.backgroundColor = this._colorInRgb;

		this._contentContainer.addEventListener('mouseover', () => {
			this._contentContainer.style.backgroundColor = "white";
			this._lookUpItem.style.backgroundColor = this._colorInRgb;
		});

		this._contentContainer.addEventListener('mouseout', () => {
			this._contentContainer.style.backgroundColor = this._colorInRgb;
		});
	}

	private performLookupObjects(entityType: string, viewId: string, setSelected: (value: ComponentFramework.LookupValue, update?: boolean) => void): void {
		// Used cached values from lookup parameter to set options for lookupObjects API
		const lookupOptions = {
			defaultEntityType: entityType,
			defaultViewId: viewId,
			allowMultiSelect: false,
			entityTypes: [entityType],
			viewIds: [viewId]
		};

		this._context.utils.lookupObjects(lookupOptions).then((success) => {
			if (success && success.length > 0) {
				// Cache the necessary information for the newly selected entity lookup
				const selectedReference = success[0];
				const selectedLookupValue: ComponentFramework.LookupValue = {
					id: selectedReference.id,
					name: selectedReference.name,
					entityType: selectedReference.entityType
				};

				// Update the primary or secondary lookup property
				setSelected(selectedLookupValue);

				// Trigger a control update
				this._notifyOutputChanged();
			} else {
				setSelected({} as ComponentFramework.LookupValue);
				this._notifyOutputChanged();
			}
		}, (error) => {
			console.log(error);
		});
	}

	private clearLookupField(): void {
		this._selectedItem = undefined;
		this._inputData.textContent = "";
		this._notifyOutputChanged();
	}

	public updateView(context: ComponentFramework.Context<IInputs>): void {
		const lookupValue: ComponentFramework.LookupValue = context.parameters.controlValue.raw[0];
		this._context = context;
		let propertyValue = lookupValue ? lookupValue.name : "";

		// Ensure propertyValue is not undefined
		if (propertyValue === undefined) {
			propertyValue = "";
		}

		if (propertyValue.trim() === "") {
			this._lookUpItem.style.display = "none"; // Hide the lookUpItem
		} else {
			this._lookUpItem.style.display = "flex"; // Show the lookUpItem
		}

		this._inputData.textContent = propertyValue;

		this._colorInRgb = context.parameters.ColorInRGB.raw || "";
		this._contentContainer.style.backgroundColor = this._colorInRgb;
		this._lookUpItem.style.backgroundColor = this._colorInRgb;
	}

	public getOutputs(): IOutputs {
		// Send the updated selected lookup item back to the ComponentFramework, based on the currently selected item
		return { controlValue: this._selectedItem ? [this._selectedItem] : [] };
	}

	public destroy(): void { }
}
