import * as React from "react";
import styles from "./Anniversaries.module.scss";
import { IAnniversariesProps } from "./IAnniversariesProps";
import { getSP } from "../../pnpjsConfig";
import { SPFI } from "@pnp/sp";
import { IAnniversariesState } from "./IAnniversariesState";
import { SearchQueryBuilder, SearchResults } from "@pnp/sp/search";
import * as moment from "moment";
import {
	Persona,
	PersonaPresence,
	IPersonaProps,
} from "@fluentui/react/lib/Persona";
import { Stack } from "@fluentui/react/lib/Stack";
import Person from "../model/Person";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { IconButton } from "@fluentui/react/lib/Button";
import { MessageBar } from "@fluentui/react";
import { escape } from "@microsoft/sp-lodash-subset";
import { IIconProps, Icon } from "@fluentui/react/lib/Icon";
import { Pagination, WebPartTitle } from "@pnp/spfx-controls-react";

export default class Anniversaries extends React.Component<
	IAnniversariesProps,
	IAnniversariesState
> {
	private _sp: SPFI;
	private readonly sourceId: string = "b09a7990-05ea-4af9-81ef-edfab16c4e31";
	//private _pageSize: number = 10;
	private _currentPageNumber = 1;
	constructor(props: IAnniversariesProps) {
		super(props);
		this._sp = getSP();
		this.state = {
			searchResult: null,
			loading: true,
		};
		this._onRenderTertiaryText = this._onRenderTertiaryText.bind(this);
		this.previousPage = this.previousPage.bind(this);
		this.nextPage = this.nextPage.bind(this);
	}
	public componentDidMount(): void {
		const selectFields: Array<string> = [
			"FirstName",
			"LastName",
			"Title",
			"PreferredName",
			"WorkEmail",
			"OfficeNumber",
			"WorkPhone",
			"MobilePhone",
			"JobTitle",
			"Department",
			"Skills",
			"PastProjects",
			"BaseOfficeLocation",
			"SPS-UserType",
			"GroupId",
		];

		this.props.textField !== null &&
			selectFields.filter((x) => x === this.props.textField).length === 0 &&
			selectFields.push(this.props.textField);

		this.props.secondaryTextField !== null &&
			selectFields.filter((x) => x === this.props.secondaryTextField).length ===
				0 &&
			selectFields.push(this.props.secondaryTextField);

		this.props.tertiaryTextField !== null &&
			selectFields.filter((x) => x === this.props.tertiaryTextField).length ===
				0 &&
			selectFields.push(this.props.tertiaryTextField);

		this.props.dateField !== null &&
			selectFields.filter((x) => x === this.props.dateField).length === 0 &&
			selectFields.push(this.props.dateField);

		let filterText = "";

		if (this.props.filterField) {
			const date1 = moment();
			const date2 = moment();
			if (this.props.dateFieldAs === "1") {
				//as 2000's date
				date1.year(2000);
				date2.year(2000);
			}

			date1.add(this.props.daysFromTodayFilter, "days");
			date2.add(this.props.daysBeforeTodayFilter, "days");

			filterText = `${this.props.filterField}<=${date1.format(
				"YYYY-MM-DD"
			)} AND ${this.props.filterField}>=${date2.format("YYYY-MM-DD")}`;
		}
		if (this.props.additionalFilterKQL)
			filterText += ` AND ${this.props.additionalFilterKQL}`;

		const q = SearchQueryBuilder()
			.text(filterText)
			.selectProperties(...selectFields)
			.sourceId(this.sourceId)
			.rowLimit(this.props.pageSize);
		this._sp
			.search(q)
			.then((results: SearchResults) => {
				this.setState({ loading: false, searchResult: results });
			})
			.catch((reason: any) => {
				console.error(reason);
				this.setState({ loading: false, searchResult: null });
			});
	}
	// Render tertiary text
	private _onRenderTertiaryText = (props: IPersonaProps): JSX.Element => {
		return (
			<div>
				<span className="ms-fontWeight-semibold" style={{ color: "#71afe5" }}>
					{props.tertiaryText}
				</span>
			</div>
		);
	};
	private nextPage(): void {
		this.setState({
			...this.state,
			loading: true,
		});
		this.state.searchResult
			.getPage(++this._currentPageNumber, this.props.pageSize)
			.then((results: SearchResults) => {
				this.setState({
					loading: false,
					searchResult: results,
				});
			})
			.catch((reason: any) => {
				console.error(reason);
				this.setState({ loading: false, searchResult: null });
			});
	}
	private previousPage(): void {
		this._currentPageNumber -= 1;
		if (this._currentPageNumber < 1) this._currentPageNumber = 1;

		this.setState({
			...this.state,
			loading: true,
		});
		this.state.searchResult
			.getPage(this._currentPageNumber, this.props.pageSize)
			.then((results: SearchResults) => {
				this.setState({
					loading: false,
					searchResult: results,
				});
			})
			.catch((reason: any) => {
				console.error(reason);
				this.setState({ loading: false, searchResult: null });
			});
	}
	// Add this method to your component class
	private handlePageChange = (page: number): void => {
		this.setState({
			...this.state,
			loading: true,
		});

		this._currentPageNumber = page;

		this.state.searchResult
			.getPage(this._currentPageNumber, this.props.pageSize)
			.then((results: SearchResults) => {
				this.setState({
					loading: false,
					searchResult: results,
				});
			})
			.catch((reason: any) => {
				console.error(reason);
				this.setState({ loading: false, searchResult: null });
			});
	};

	public render(): React.ReactElement<IAnniversariesProps> {
		let peopleList: Array<Person> = [];
		const forwardIcon: IIconProps = { iconName: "ChevronRight" };
		const backIcon: IIconProps = { iconName: "ChevronLeft" };

		if (this.state.searchResult)
			peopleList = this.state.searchResult.PrimarySearchResults.map(
				(item) =>
					new Person(
						this.props.webUrl,
						item,
						this.props.textField,
						this.props.secondaryTextField,
						this.props.tertiaryTextField,
						this.props.dateField
					)
			);

		return (
			<div className={`${styles.anniversaries} `}>
				<WebPartTitle
					displayMode={this.props.displayMode}
					title={this.props.title}
					updateProperty={this.props.onTitleUpdate}
				/>
				<Stack tokens={{ childrenGap: 10 }}>
					{this.state.loading ? (
						<Spinner size={SpinnerSize.medium} />
					) : (
						peopleList.map((item, index) => (
							<div className={`${styles.birthdayCard}`} key={index}>
								{item.Date && (
									<h3>
										{moment(item.Date).format("Do [of] MMMM")}{" "}
										<Icon iconName={this.props.celebrateIcon} />
									</h3>
								)}
								<Persona
									imageUrl={item.PictureURL}
									text={item.Text}
									tertiaryText={item.TertiaryText}
									secondaryText={item.SecondaryText}
									size={this.props.personaSize}
									presence={PersonaPresence.none}
									showSecondaryText={true}
								/>
							</div>
						))
					)}
					{!this.state.loading && this.state.searchResult.TotalRows === 0 && (
						<MessageBar>{escape(this.props.noResultsMessage)}</MessageBar>
					)}
					{this.state.searchResult !== null && !this.state.loading && (
						<Stack
							horizontal
							verticalAlign="center"
							tokens={{ childrenGap: 10 }}
							className={styles.pagination_buttons}
						>
							<IconButton
								disabled={this._currentPageNumber === 1}
								onClick={this.previousPage}
								iconProps={backIcon}
							/>
							<div className={styles.pagination}>
								<Pagination
									currentPage={this._currentPageNumber}
									totalPages={Math.ceil(
										this.state.searchResult.TotalRows / this.props.pageSize
									)}
									onChange={(page: number) => this.handlePageChange(page)}
									hideFirstPageJump
									hideLastPageJump
								/>
							</div>
							<IconButton
								disabled={
									!(
										Math.ceil(
											this.state.searchResult.TotalRows / this.props.pageSize
										) > this._currentPageNumber
									)
								}
								onClick={this.nextPage}
								iconProps={forwardIcon}
							/>
						</Stack>
					)}
				</Stack>
			</div>
		);
	}
}
