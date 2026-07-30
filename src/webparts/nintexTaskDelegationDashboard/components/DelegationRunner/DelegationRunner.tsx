import * as React from 'react';
import { IDelegationRunnerProps } from './IDelegationRunnerProps';
import { 
  NormalPeoplePicker, 
  IPersonaProps, 
  PrimaryButton, 
  DefaultButton, 
  ProgressIndicator, 
  MessageBar, 
  MessageBarType, 
  DetailsList, 
  IColumn, 
  Dialog, 
  DialogType, 
  DialogFooter, 
  Icon, 
  IconButton, 
  TextField, 
  DatePicker, 
  Spinner, 
  SpinnerSize,
  ContextualMenu,
  IContextualMenuItem,
  DirectionalHint,
  Persona,
  PersonaSize,
  TooltipHost
} from '@fluentui/react';
import { TokenService } from '../../../../services/TokenService';
import { NintexApiService, INintexAutoDelegation, INintexUser } from '../../../../services/NintexApiService';
import { SPHttpClient } from '@microsoft/sp-http';

export type DelegationStatusFilter = 'Active' | 'Scheduled' | 'Expired' | 'All';

interface IOverflowTextProps {
  text: string;
  style?: React.CSSProperties;
  className?: string;
}

const OverflowText: React.FC<IOverflowTextProps> = ({ text, style, className }) => {
  const [isOverflowing, setIsOverflowing] = React.useState<boolean>(false);
  const spanRef = React.useRef<HTMLSpanElement>(null);

  const checkOverflow = (): void => {
    if (spanRef.current) {
      const hasOverflow = spanRef.current.scrollWidth > spanRef.current.clientWidth;
      if (hasOverflow !== isOverflowing) {
        setIsOverflowing(hasOverflow);
      }
    }
  };

  return (
    <TooltipHost
      content={isOverflowing ? text : ''}
      styles={{ root: { display: 'block', overflow: 'hidden', minWidth: 0 } }}
    >
      <span
        ref={spanRef}
        onMouseEnter={checkOverflow}
        style={{
          display: 'block',
          whiteSpace: 'nowrap',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          minWidth: 0,
          ...style
        }}
        className={className}
      >
        {text}
      </span>
    </TooltipHost>
  );
};

export const DelegationRunner: React.FC<IDelegationRunnerProps> = (props) => {
  const [delegateUser, setDelegateUser] = React.useState<IPersonaProps | undefined>(undefined);
  const [oooUser, setOooUser] = React.useState<IPersonaProps | undefined>(undefined);
  const [nintexToken, setNintexToken] = React.useState<string>("");
  const [delegations, setDelegations] = React.useState<INintexAutoDelegation[]>([]);
  const [isDelegating, setIsDelegating] = React.useState<boolean>(false);
  const [isPanelOpen, setIsPanelOpen] = React.useState<boolean>(false);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [dateFrom, setDateFrom] = React.useState<Date | undefined>(undefined);
  const [timeFrom, setTimeFrom] = React.useState<string>("00:00");
  const [dateTo, setDateTo] = React.useState<Date | undefined>(undefined);
  const [timeTo, setTimeTo] = React.useState<string>("00:00");
  const [message, setMessage] = React.useState<string>("");
  const [progress, setProgress] = React.useState<number>(0);
  const [progressDescription, setProgressDescription] = React.useState<string>("");
  const [errorMsg, setErrorMsg] = React.useState<string>("");
  const [successMsg, setSuccessMsg] = React.useState<string>("");
  const [editingId, setEditingId] = React.useState<string | null>(null);
  const [initialFormState, setInitialFormState] = React.useState<{ dateFrom: string; timeFrom: string; dateTo: string; timeTo: string; message: string } | null>(null);
  const [isDeleteDialogOpen, setIsDeleteDialogOpen] = React.useState<boolean>(false);
  const [itemToDelete, setItemToDelete] = React.useState<INintexAutoDelegation | null>(null);

  const userCacheRef = React.useRef<{ [userId: string]: INintexUser }>({});

  // Status Filter and Pagination State
  const [statusFilter, setStatusFilter] = React.useState<DelegationStatusFilter>('Active');
  const [currentPage, setCurrentPage] = React.useState<number>(1);
  const pageSize = 20;

  // Contextual menu state for column filter
  const [contextualMenuTarget, setContextualMenuTarget] = React.useState<MouseEvent | HTMLElement | undefined>(undefined);
  const [showContextMenu, setShowContextMenu] = React.useState<boolean>(false);

  // User detail dialog state
  const [isUserDetailOpen, setIsUserDetailOpen] = React.useState<boolean>(false);
  const [userDetailLoading, setUserDetailLoading] = React.useState<boolean>(false);
  const [userDetailError, setUserDetailError] = React.useState<string>("");
  const [selectedNintexUser, setSelectedNintexUser] = React.useState<INintexUser | null>(null);
  const [selectedUserId, setSelectedUserId] = React.useState<string>("");

  const handleGetUserDetails = async (userId: string): Promise<void> => {
    setSelectedUserId(userId);
    setSelectedNintexUser(null);
    setUserDetailError("");
    setUserDetailLoading(true);
    setIsUserDetailOpen(true);

    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const nintexApiService = new NintexApiService(props.context.httpClient as any, props.nintexApiBaseUrl);
      const user = await nintexApiService.getNintexUserById(userId, nintexToken);
      if (user) {
        if (user.id) {
          userCacheRef.current[user.id] = user;
        }
        setSelectedNintexUser(user);
      } else {
        setUserDetailError(`No user details found for ID: ${userId}`);
      }
    } catch (err) {
      console.error("Error fetching user details:", err);
      setUserDetailError(err.message || "Failed to fetch user details.");
    } finally {
      setUserDetailLoading(false);
    }
  };

  const loadDelegations = async (apiService: NintexApiService, token: string, status: DelegationStatusFilter = 'Active'): Promise<void> => {
    setIsLoading(true);
    try {
      const allDelegations = await apiService.listAutoDelegations(token, status);
      
      // Set fromUserDisplay to email (userId) directly without user profile API calls
      const processed = allDelegations.map((d: INintexAutoDelegation) => ({
        ...d,
        fromUserDisplay: d.userId
      }));

      setDelegations(processed);
    } finally {
      setIsLoading(false);
    }
  };

  React.useEffect(() => {
    const fetchToken = async (): Promise<void> => {
      try {
        setIsLoading(true);
        const spHttpClient: SPHttpClient = props.context.spHttpClient;
        const tokenService = new TokenService(spHttpClient, props.tokenListUrl, props.tokenTitleValue, props.tokenColumnName);
        const token = await tokenService.getToken();
        setNintexToken(token);
        
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const nintexApiService = new NintexApiService(props.context.httpClient as any, props.nintexApiBaseUrl);
        await loadDelegations(nintexApiService, token, 'Active');
      } catch (err) {
        console.error("Error fetching initial token:", err);
        setErrorMsg("Failed to initialize Nintex API token.");
      } finally {
        setIsLoading(false);
      }
    };
    fetchToken().catch(console.error);
  }, [props.tokenListUrl, props.tokenTitleValue, props.tokenColumnName]);

  const onResolveSuggestions = async (filterText: string, ignoreIds: (string | undefined)[] = []): Promise<IPersonaProps[]> => {
    if (!filterText || filterText.length < 2 || !nintexToken) return [];
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const nintexApiService = new NintexApiService(props.context.httpClient as any, props.nintexApiBaseUrl);
      const users = await nintexApiService.searchNintexUsers(filterText, nintexToken);
      users.forEach((u: INintexUser) => {
        if (u.id) {
          userCacheRef.current[u.id] = u;
        }
      });
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return users
        .filter((u: INintexUser) => ignoreIds.indexOf(u.id) === -1)
        .map((u: INintexUser) => ({
          text: `${u.firstName} ${u.lastName}`,
          secondaryText: u.email,
          id: u.id,
          imageUrl: undefined
        }));
    } catch (e) {
      console.error("Error fetching Nintex users:", e);
      return [];
    }
  };

  const handleEditClick = (item: INintexAutoDelegation): void => {
    setEditingId(item.id);
    let initDateFromStr = "";
    let initTimeFromStr = "00:00";
    if (item.effectiveFrom) {
      const d = new Date(item.effectiveFrom);
      initDateFromStr = d.toDateString();
      initTimeFromStr = `${('0' + d.getHours()).slice(-2)}:${('0' + d.getMinutes()).slice(-2)}`;
      setDateFrom(d);
      setTimeFrom(initTimeFromStr);
    } else {
      setDateFrom(undefined);
      setTimeFrom("00:00");
    }
    
    let initDateToStr = "";
    let initTimeToStr = "00:00";
    if (item.effectiveTo) {
      const d = new Date(item.effectiveTo);
      initDateToStr = d.toDateString();
      initTimeToStr = `${('0' + d.getHours()).slice(-2)}:${('0' + d.getMinutes()).slice(-2)}`;
      setDateTo(d);
      setTimeTo(initTimeToStr);
    } else {
      setDateTo(undefined);
      setTimeTo("00:00");
    }
    const initMsg = item.message || '';
    setMessage(initMsg);

    setInitialFormState({
      dateFrom: initDateFromStr,
      timeFrom: initTimeFromStr,
      dateTo: initDateToStr,
      timeTo: initTimeToStr,
      message: initMsg
    });

    // Display Email - resolve on demand if userId is a Nintex user ID
    if (item.userId) {
      const isEmail = item.userId.indexOf('@') !== -1;
      const cachedUser = userCacheRef.current[item.userId];
      const initialDisplayEmail = isEmail ? item.userId : (cachedUser?.email || item.userId);

      setOooUser({
        id: item.userId,
        text: initialDisplayEmail,
        secondaryText: initialDisplayEmail
      });

      // If it's not an email and not yet cached, resolve the user email on demand
      if (!isEmail && !cachedUser && nintexToken) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const nintexApiService = new NintexApiService(props.context.httpClient as any, props.nintexApiBaseUrl);
        nintexApiService.getNintexUserById(item.userId, nintexToken).then(user => {
          if (user && user.email) {
            userCacheRef.current[item.userId] = user;
            setOooUser({
              id: item.userId,
              text: user.email,
              secondaryText: user.email
            });
          }
        }).catch(err => {
          console.error("Error resolving user email for edit dialog:", err);
        });
      }
    } else {
      setOooUser(undefined);
    }

    if (item.standIns && item.standIns.length > 0) {
      const standIn = item.standIns[0];
      const delegateEmail = standIn.emails?.[0] || standIn.id;
      setDelegateUser({
        id: standIn.id,
        text: delegateEmail,
        secondaryText: delegateEmail
      });
    } else {
      setDelegateUser(undefined);
    }
    
    setErrorMsg("");
    setSuccessMsg("");
    setIsPanelOpen(true);
  };

  const handleDeleteClick = (item: INintexAutoDelegation): void => {
    setItemToDelete(item);
    setIsDeleteDialogOpen(true);

    if (item.userId && item.userId.indexOf('@') === -1 && !userCacheRef.current[item.userId] && nintexToken) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const nintexApiService = new NintexApiService(props.context.httpClient as any, props.nintexApiBaseUrl);
      nintexApiService.getNintexUserById(item.userId, nintexToken).then(user => {
        if (user && user.email && user.id) {
          const currentCache = userCacheRef.current;
          currentCache[user.id] = user;
        }
      }).catch(console.error);
    }
  };

  const confirmDelete = async (): Promise<void> => {
    if (!itemToDelete) return;
    setIsDelegating(true);
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const nintexApiService = new NintexApiService(props.context.httpClient as any, props.nintexApiBaseUrl);
      await nintexApiService.deleteAutoDelegation(itemToDelete.id, nintexToken);
      await loadDelegations(nintexApiService, nintexToken, statusFilter);
      setIsDeleteDialogOpen(false);
      setItemToDelete(null);
    } catch (err) {
      alert(`Failed to delete: ${err.message}`);
    } finally {
      setIsDelegating(false);
    }
  };

  const handleDelegate = async (): Promise<void> => {
    setErrorMsg("");
    setSuccessMsg("");
    
    if (!oooUser || !oooUser.id) {
      setErrorMsg("Please select a From User (Delegator).");
      return;
    }
    if (!delegateUser || !delegateUser.id) {
      setErrorMsg("Please select a Delegate User.");
      return;
    }
    if (oooUser.id === delegateUser.id || 
        (oooUser.secondaryText && delegateUser.secondaryText && 
         oooUser.secondaryText.toLowerCase() === delegateUser.secondaryText.toLowerCase())) {
      setErrorMsg("Delegate from and Delegate to users cannot be the same.");
      return;
    }
    if (!dateFrom || !dateTo || !timeFrom || !timeTo) {
      setErrorMsg("Please select both Start and End date and time.");
      return;
    }

    const fromDateObj = new Date(dateFrom.getTime());
    const [fromH, fromM] = timeFrom.split(':').map(Number);
    fromDateObj.setHours(fromH, fromM, 0, 0);

    const toDateObj = new Date(dateTo.getTime());
    const [toH, toM] = timeTo.split(':').map(Number);
    toDateObj.setHours(toH, toM, 0, 0);
    
    const now = new Date();
    now.setSeconds(0, 0);

    if (toDateObj <= fromDateObj) {
      setErrorMsg("End date and time must be after Start date and time.");
      return;
    }
    
    if (toDateObj < now) {
      setErrorMsg("End date and time cannot be in the past.");
      return;
    }

    if (!editingId && fromDateObj < now) {
      setErrorMsg("Start date and time cannot be in the past.");
      return;
    }

    setIsDelegating(true);
    setProgress(0);
    setProgressDescription("Configuring auto-delegation rule...");

    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const nintexApiService = new NintexApiService(props.context.httpClient as any, props.nintexApiBaseUrl);

      const delegatorUserId = (oooUser.id || oooUser.secondaryText) as string;
      const delegatorEmail = (oooUser.secondaryText || oooUser.id) as string;
      const delegateEmail = (delegateUser.secondaryText || delegateUser.id) as string;

      if (!delegatorUserId || !delegatorEmail) {
        setErrorMsg("Could not determine the 'Delegate from' user. Please re-select the user.");
        setIsDelegating(false);
        return;
      }
      if (!delegateEmail) {
        setErrorMsg("Could not determine the email address for the 'Delegate to' user. Please re-select the user.");
        setIsDelegating(false);
        return;
      }

      if (editingId) {
        setProgressDescription("Updating auto-delegation rule...");
        // Pass delegatorUserId (actual Nintex User ID, e.g. auth0|...) for update payload
        await nintexApiService.updateAutoDelegation(editingId, delegatorUserId, delegateEmail, fromDateObj, toDateObj, nintexToken, message);
        setSuccessMsg(`Successfully updated auto-delegation rule!`);
      } else {
        setProgressDescription("Configuring auto-delegation rule...");
        await nintexApiService.createAutoDelegation(delegatorEmail, delegateEmail, fromDateObj, toDateObj, nintexToken, message);
        setSuccessMsg(`Successfully created auto-delegation rule!`);
      }
      
      setProgress(1);
      setProgressDescription("Completed!");
      
      await loadDelegations(nintexApiService, nintexToken, statusFilter);

      setIsPanelOpen(false);
      setEditingId(null);
      setSuccessMsg("");
      setDelegateUser(undefined);
      setOooUser(undefined);
      setDateFrom(undefined);
      setTimeFrom("00:00");
      setDateTo(undefined);
      setTimeTo("00:00");
      setMessage("");
      setProgress(0);
    } catch (err) {
      const errMsg: string = err.message || "";
      if (errMsg.indexOf("UserNotFound") !== -1 || errMsg.indexOf("is not found") !== -1) {
        setErrorMsg(
          "One of the selected users has not been provisioned in the Nintex workflow tasks system. " +
          "The user must interact with Nintex (e.g. be assigned a task) before they can be used for auto-delegation. " +
          "Please contact your Nintex administrator."
        );
      } else {
        setErrorMsg(errMsg || "An unexpected error occurred.");
      }

    } finally {
      setIsDelegating(false);
    }
  };

  const delegationColumns: IColumn[] = [
    { 
      key: 'col0', 
      name: 'Delegate from', 
      fieldName: 'userId', 
      minWidth: 160, 
      maxWidth: 220, 
      onRender: (item: INintexAutoDelegation) => {
        return (
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', width: '100%' }}>
            <div style={{ display: 'flex', alignItems: 'center', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              <Icon iconName="Contact" style={{ marginRight: '8px', flexShrink: 0 }} />
              <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }} title={item.userId}>{item.userId}</span>
            </div>
            <TooltipHost content="Get user details?">
              <IconButton
                iconProps={{ iconName: 'Info' }}
                title="Get user details"
                ariaLabel="Get user details"
                onClick={(e: React.MouseEvent<HTMLButtonElement | HTMLAnchorElement>) => {
                  e.stopPropagation();
                  handleGetUserDetails(item.userId).catch(console.error);
                }}
                styles={{
                  root: { height: '24px', width: '24px', marginLeft: '4px' }
                }}
              />
            </TooltipHost>
          </div>
        );
      }
    },
    { 
      key: 'col3', 
      name: 'Delegate to', 
      fieldName: 'standIns', 
      minWidth: 160, 
      maxWidth: 220, 
      onRender: (item) => {
        const standIn = item.standIns?.[0];
        const displayName = standIn ? (standIn.emails?.[0] || standIn.id) : '';
        return (
          <div style={{ display: 'flex', alignItems: 'center' }}>
            <Icon iconName="Contact" style={{ marginRight: '8px' }} />
            <span>{displayName}</span>
          </div>
        );
      }
    },
    { 
      key: 'col1', 
      name: 'Effective From', 
      fieldName: 'effectiveFrom', 
      minWidth: 130, 
      maxWidth: 150, 
      onRender: (item) => {
        const d = new Date(item.effectiveFrom);
        return `${d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}, ${d.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })}`;
      }
    },
    { 
      key: 'col2', 
      name: 'Effective To', 
      fieldName: 'effectiveTo', 
      minWidth: 130, 
      maxWidth: 150, 
      onRender: (item) => {
        const d = new Date(item.effectiveTo);
        return `${d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}, ${d.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })}`;
      }
    },

    { 
      key: 'col4', 
      name: 'Message', 
      fieldName: 'message', 
      minWidth: 180, 
      maxWidth: 280 
    },
    { 
      key: 'col5', 
      name: 'Status', 
      fieldName: 'status', 
      minWidth: 90, 
      maxWidth: 110, 
      isFiltered: true,
      onColumnClick: (ev?: React.MouseEvent<HTMLElement>) => {
        if (ev) {
          setContextualMenuTarget(ev.currentTarget);
          setShowContextMenu(true);
        }
      },
      onRender: (item) => {
        const now = new Date();
        const from = new Date(item.effectiveFrom);
        const to = new Date(item.effectiveTo);
        const isActive = now >= from && now <= to;
        const text = isActive ? 'Active' : (now < from ? 'Scheduled' : 'Expired');
        const color = isActive ? '#107c10' : (text === 'Scheduled' ? '#0078d4' : '#605e5c');
        const bgColor = 'transparent';
        
        return (
          <span style={{ 
            padding: '2px 8px', 
            borderRadius: '12px', 
            border: `1px solid ${color}`, 
            color: color, 
            backgroundColor: bgColor,
            fontSize: '12px'
          }}>
            {text}
          </span>
        );
      }
    },

    {
      key: 'col6',
      name: '',
      fieldName: 'actions',
      minWidth: 40,
      maxWidth: 40,
      onRender: (item) => (
        <IconButton 
          menuProps={{
            items: [
              {
                key: 'edit',
                text: 'Edit',
                iconProps: { iconName: 'Edit' },
                onClick: () => handleEditClick(item)
              },
              {
                key: 'delete',
                text: 'Delete',
                iconProps: { iconName: 'Delete' },
                onClick: () => handleDeleteClick(item)
              }
            ]
          }}
          iconProps={{ iconName: 'More' }} 
          title="Actions" 
          ariaLabel="Actions" 
        />
      )
    }
  ];

  // Data is already filtered by API service
  const filteredDelegations = delegations;

  const totalPages = Math.ceil(filteredDelegations.length / pageSize) || 1;
  const validPage = Math.min(Math.max(currentPage, 1), totalPages);
  const pagedItems = filteredDelegations.slice((validPage - 1) * pageSize, validPage * pageSize);

  const bufferMs = (props.bufferMinutes !== undefined ? props.bufferMinutes : 5) * 60000;
  const now = new Date();
  const earliestAllowed = new Date(now.getTime() + bufferMs);
  
  const earliestAllowedDate = new Date(earliestAllowed.getTime());
  earliestAllowedDate.setHours(0, 0, 0, 0);

  let minStartDate = earliestAllowedDate;
  if (editingId && dateFrom && dateFrom < earliestAllowedDate) {
    minStartDate = new Date(dateFrom.getTime());
    minStartDate.setHours(0, 0, 0, 0);
  }

  let minEndDate = dateFrom ? new Date(dateFrom.getTime()) : earliestAllowedDate;
  minEndDate.setHours(0, 0, 0, 0);
  if (minEndDate < earliestAllowedDate) {
    minEndDate = earliestAllowedDate;
  }

  const currentTimeStr = `${('0' + now.getHours()).slice(-2)}:${('0' + now.getMinutes()).slice(-2)}`;
  const bufferTimeStr = `${('0' + earliestAllowed.getHours()).slice(-2)}:${('0' + earliestAllowed.getMinutes()).slice(-2)}`;
  
  const isStartDateEarliest = dateFrom && dateFrom.toDateString() === earliestAllowed.toDateString();
  const startMinTime = (!editingId && isStartDateEarliest) ? bufferTimeStr : undefined;

  const isEndDateToday = dateTo && dateTo.toDateString() === new Date().toDateString();
  const isEndDateSameAsStartDate = dateTo && dateFrom && dateTo.toDateString() === dateFrom.toDateString();
  
  let endMinTime: string | undefined = undefined;
  if (isEndDateToday) {
    endMinTime = currentTimeStr;
  }
  if (isEndDateSameAsStartDate) {
    endMinTime = endMinTime && endMinTime > timeFrom ? endMinTime : timeFrom;
  }

  const currentDateFromStr = dateFrom ? dateFrom.toDateString() : '';
  const currentDateToStr = dateTo ? dateTo.toDateString() : '';
  const currentMsg = message || '';

  const hasFormChanged = !editingId || (
    initialFormState !== null && (
      currentDateFromStr !== initialFormState.dateFrom ||
      timeFrom !== initialFormState.timeFrom ||
      currentDateToStr !== initialFormState.dateTo ||
      timeTo !== initialFormState.timeTo ||
      currentMsg !== initialFormState.message
    )
  );

  const onFilterChange = async (newStatus: DelegationStatusFilter): Promise<void> => {
    setStatusFilter(newStatus);
    setCurrentPage(1);
    setShowContextMenu(false);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const nintexApiService = new NintexApiService(props.context.httpClient as any, props.nintexApiBaseUrl);
    await loadDelegations(nintexApiService, nintexToken, newStatus);
  };

  const menuItems: IContextualMenuItem[] = [
    { key: 'Active', text: 'Active', onClick: () => { onFilterChange('Active').catch(console.error); }, canCheck: true, isChecked: statusFilter === 'Active' },
    { key: 'Scheduled', text: 'Scheduled', onClick: () => { onFilterChange('Scheduled').catch(console.error); }, canCheck: true, isChecked: statusFilter === 'Scheduled' },
    { key: 'Expired', text: 'Expired (Last 30 Days)', onClick: () => { onFilterChange('Expired').catch(console.error); }, canCheck: true, isChecked: statusFilter === 'Expired' },
    { key: 'All', text: 'All', onClick: () => { onFilterChange('All').catch(console.error); }, canCheck: true, isChecked: statusFilter === 'All' }
  ];

  return (
    <div style={{ padding: '20px', backgroundColor: '#fff', boxShadow: '0 2px 4px rgba(0,0,0,0.1)' }}>
      {showContextMenu && (
        <ContextualMenu
          items={menuItems}
          hidden={!showContextMenu}
          target={contextualMenuTarget}
          onItemClick={() => setShowContextMenu(false)}
          onDismiss={() => setShowContextMenu(false)}
          directionalHint={DirectionalHint.bottomLeftEdge}
        />
      )}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px', flexWrap: 'wrap', gap: '10px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          <span style={{ fontSize: '13px', color: '#605e5c' }}>Filter:</span>
          <DefaultButton
            text={`Status: ${statusFilter === 'Expired' ? 'Expired (Last 30 Days)' : statusFilter}`}
            iconProps={{ iconName: 'Filter' }}
            onClick={(ev: React.MouseEvent<HTMLElement>) => {
              setContextualMenuTarget(ev.currentTarget);
              setShowContextMenu(true);
            }}
          />
        </div>
        <PrimaryButton 
          text="Add Auto Task Delegation" 
          onClick={() => {
            setEditingId(null);
            setInitialFormState(null);
            setDelegateUser(undefined);
            setOooUser(undefined);
            setDateFrom(undefined);
            setTimeFrom("00:00");
            setDateTo(undefined);
            setTimeTo("00:00");
            setMessage("");
            setErrorMsg("");
            setSuccessMsg("");
            setIsPanelOpen(true);
          }} 
          disabled={!nintexToken} 
        />
      </div>

      <div>
        {isLoading ? (
          <div style={{ padding: '40px 0', display: 'flex', justifyContent: 'center' }}>
            <Spinner size={SpinnerSize.large} label="Loading delegations..." />
          </div>
        ) : (
          <>
            <DetailsList
              items={pagedItems}
              columns={delegationColumns}
              setKey="set"
              selectionMode={0}
            />
            {filteredDelegations.length === 0 && (
              <div style={{ padding: '30px', textAlign: 'center', color: '#605e5c', backgroundColor: '#faf9f8', borderBottom: '1px solid #edebe9' }}>
                No {statusFilter.toLowerCase()} auto-delegations found.
              </div>
            )}
            {filteredDelegations.length > 0 && (
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '15px', padding: '10px 0', borderTop: '1px solid #edebe9' }}>
                <span style={{ fontSize: '13px', color: '#605e5c' }}>
                  Showing {(validPage - 1) * pageSize + 1} - {Math.min(validPage * pageSize, filteredDelegations.length)} of {filteredDelegations.length} items
                </span>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <IconButton
                    iconProps={{ iconName: 'ChevronLeft' }}
                    title="Previous Page"
                    ariaLabel="Previous Page"
                    disabled={validPage <= 1}
                    onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
                  />
                  <span style={{ fontSize: '13px', fontWeight: 600 }}>
                    Page {validPage} of {totalPages}
                  </span>
                  <IconButton
                    iconProps={{ iconName: 'ChevronRight' }}
                    title="Next Page"
                    ariaLabel="Next Page"
                    disabled={validPage >= totalPages}
                    onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
                  />
                </div>
              </div>
            )}
          </>
        )}
      </div>

      <Dialog
        hidden={!isPanelOpen}
        onDismiss={() => {
          setIsPanelOpen(false);
          setEditingId(null);
          setErrorMsg("");
          setSuccessMsg("");
        }}
        dialogContentProps={{
          type: DialogType.normal,
          title: <span style={{ color: '#d83b01' }}>{editingId ? "Edit" : "Add"} auto task delegation</span>,
          showCloseButton: true
        }}
        modalProps={{
          isBlocking: false,
          styles: { main: { maxWidth: 700, width: '100%' } }
        }}
      >
        <div style={{ display: 'flex', flexDirection: 'column', gap: '20px', padding: '0 5px' }}>
          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap' }}>
            <div style={{ flex: '1 1 200px' }}>
              <label style={{ display: 'block', fontWeight: 600, fontSize: '12px', marginBottom: '2px' }}>Delegate from</label>
              <NormalPeoplePicker
                onResolveSuggestions={(filterText) => onResolveSuggestions(filterText, [delegateUser?.id])}
                itemLimit={1}
                disabled={isDelegating || !nintexToken || !!editingId}
                onChange={(items) => setOooUser(items && items.length > 0 ? items[0] : undefined)}
                selectedItems={oooUser ? [oooUser] : []}
                resolveDelay={500}
                styles={{ root: { maxWidth: '100%' } }}
                inputProps={{ placeholder: "Search for and select 1 Nintex user" }}
              />
            </div>

            <div style={{ flex: '1 1 200px' }}>
              <label style={{ display: 'block', fontWeight: 600, fontSize: '12px', marginBottom: '2px' }}>Delegate to</label>
              <NormalPeoplePicker
                onResolveSuggestions={(filterText) => onResolveSuggestions(filterText, [oooUser?.id])}
                itemLimit={1}
                disabled={isDelegating || !nintexToken || !!editingId}
                onChange={(items) => setDelegateUser(items && items.length > 0 ? items[0] : undefined)}
                selectedItems={delegateUser ? [delegateUser] : []}
                resolveDelay={500}
                styles={{ root: { maxWidth: '100%' } }}
                inputProps={{ placeholder: "Search for and select 1 Nintex user" }}
              />
            </div>
          </div>

          {editingId && (
            <MessageBar messageBarType={MessageBarType.info}>
              Delegator and Delegatee cannot be changed on an existing rule. Only dates and messages can be updated.
            </MessageBar>
          )}

          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap' }}>
            <div style={{ flex: '1 1 250px' }}>
              <div style={{ display: 'flex', gap: '10px' }}>
                <div style={{ flex: 1 }}>
                  <DatePicker
                    label="Start date"
                    value={dateFrom}
                    onSelectDate={(date: Date | null | undefined) => {
                      setDateFrom(date || undefined);
                      if (date && !editingId && date.toDateString() === earliestAllowed.toDateString()) {
                        if (timeFrom < bufferTimeStr) setTimeFrom(bufferTimeStr);
                      }
                    }}
                    disabled={isDelegating}
                    allowTextInput
                    minDate={minStartDate}
                    styles={{ root: { width: '100%' } }}
                  />
                </div>
                <div style={{ width: '130px' }}>
                  <TextField 
                    label="Start time" 
                    type="time" 
                    value={timeFrom} 
                    onChange={(e, val) => {
                      let newVal = val || "00:00";
                      if (startMinTime && newVal < startMinTime) {
                        newVal = startMinTime;
                      }
                      setTimeFrom(newVal);
                      if (isEndDateSameAsStartDate && timeTo < newVal) {
                        setTimeTo(newVal);
                      }
                    }}
                    disabled={isDelegating}
                    min={startMinTime}
                  />
                </div>
              </div>
            </div>
            <div style={{ flex: '1 1 250px' }}>
              <div style={{ display: 'flex', gap: '10px' }}>
                <div style={{ flex: 1 }}>
                  <DatePicker
                    label="End date"
                    value={dateTo}
                    onSelectDate={(date: Date | null | undefined) => {
                      setDateTo(date || undefined);
                      if (date) {
                        let eMin = undefined;
                        const nowStr = `${('0' + new Date().getHours()).slice(-2)}:${('0' + new Date().getMinutes()).slice(-2)}`;
                        if (date.toDateString() === new Date().toDateString()) eMin = nowStr;
                        if (dateFrom && date.toDateString() === dateFrom.toDateString()) {
                          eMin = eMin && eMin > timeFrom ? eMin : timeFrom;
                        }
                        if (eMin && timeTo < eMin) setTimeTo(eMin);
                      }
                    }}
                    disabled={isDelegating}
                    allowTextInput
                    minDate={minEndDate}
                    styles={{ root: { width: '100%' } }}
                  />
                </div>
                <div style={{ width: '130px' }}>
                  <TextField 
                    label="End time" 
                    type="time" 
                    value={timeTo} 
                    onChange={(e, val) => {
                      let newVal = val || "00:00";
                      if (endMinTime && newVal < endMinTime) {
                        newVal = endMinTime;
                      }
                      setTimeTo(newVal);
                    }}
                    disabled={isDelegating}
                    min={endMinTime}
                  />
                </div>
              </div>
            </div>
          </div>

          <div>
            <label style={{ display: 'block', fontWeight: 600, fontSize: '12px', marginBottom: '5px' }}>Message to delegate (optional)</label>
            <TextField 
              multiline 
              rows={3} 
              value={message}
              onChange={(e, val) => setMessage(val || '')}
              disabled={isDelegating}
            />
          </div>

          {isDelegating && (
            <ProgressIndicator label="Delegation Progress" description={progressDescription} percentComplete={progress} />
          )}

          {errorMsg && (
            <MessageBar messageBarType={MessageBarType.error}>{errorMsg}</MessageBar>
          )}

          {successMsg && (
            <MessageBar messageBarType={MessageBarType.success}>{successMsg}</MessageBar>
          )}
        </div>
        <DialogFooter>
          {isDelegating && <span style={{ display: 'inline-block', verticalAlign: 'middle', marginRight: '10px' }}><Spinner size={SpinnerSize.small} /></span>}
          <DefaultButton onClick={() => {
            setIsPanelOpen(false);
            setEditingId(null);
            setInitialFormState(null);
            setErrorMsg("");
            setSuccessMsg("");
          }} text="Cancel" disabled={isDelegating} />
          {hasFormChanged && (
            <PrimaryButton onClick={handleDelegate} text={editingId ? "Update" : "Add"} disabled={isDelegating} />
          )}
        </DialogFooter>
      </Dialog>

      <Dialog
        hidden={!isDeleteDialogOpen}
        onDismiss={() => setIsDeleteDialogOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Delete Task Delegation',
          subText: 'Are you sure you want to delete this task delegation? This action cannot be undone.'
        }}
        modalProps={{
          isBlocking: false,
          styles: { main: { minWidth: 480, maxWidth: 540 } }
        }}
      >
        {itemToDelete && (
          <div style={{
            marginTop: '12px',
            backgroundColor: '#f3f2f1',
            padding: '12px 16px',
            borderRadius: '4px',
            fontSize: '13px',
            display: 'flex',
            flexDirection: 'column',
            gap: '6px'
          }}>
            <div style={{ display: 'grid', gridTemplateColumns: '110px 1fr', gap: '6px 12px', alignItems: 'center' }}>
              <strong>Delegate from:</strong>
              <OverflowText text={userCacheRef.current[itemToDelete.userId]?.email || itemToDelete.userId} />

              <strong>Delegate to:</strong>
              <OverflowText text={itemToDelete.standIns?.[0]?.emails?.[0] || itemToDelete.standIns?.[0]?.id || '-'} />

              <strong>Effective From:</strong>
              <span>
                {itemToDelete.effectiveFrom ? `${new Date(itemToDelete.effectiveFrom).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}, ${new Date(itemToDelete.effectiveFrom).toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })}` : '-'}
              </span>

              <strong>Effective To:</strong>
              <span>
                {itemToDelete.effectiveTo ? `${new Date(itemToDelete.effectiveTo).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}, ${new Date(itemToDelete.effectiveTo).toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })}` : '-'}
              </span>

              {itemToDelete.message ? (
                <>
                  <strong>Message:</strong>
                  <OverflowText text={itemToDelete.message} />
                </>
              ) : null}
            </div>
          </div>
        )}
        <DialogFooter>
          {isDelegating && <span style={{ display: 'inline-block', verticalAlign: 'middle', marginRight: '10px' }}><Spinner size={SpinnerSize.small} label="Deleting..." labelPosition="left" /></span>}
          <PrimaryButton onClick={confirmDelete} text="Delete" disabled={isDelegating} style={{ backgroundColor: '#d13438', borderColor: '#d13438' }} />
          <DefaultButton onClick={() => setIsDeleteDialogOpen(false)} text="Cancel" disabled={isDelegating} />
        </DialogFooter>
      </Dialog>

      <Dialog
        hidden={!isUserDetailOpen}
        onDismiss={() => setIsUserDetailOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Nintex User Details',
          showCloseButton: true
        }}
        modalProps={{
          isBlocking: false,
          styles: { main: { minWidth: 650, maxWidth: 750 } }
        }}
      >
        {userDetailLoading ? (
          <div style={{ padding: '20px 0', textAlign: 'center' }}>
            <Spinner size={SpinnerSize.large} label="Loading user details..." />
          </div>
        ) : userDetailError ? (
          <MessageBar messageBarType={MessageBarType.error}>{userDetailError}</MessageBar>
        ) : selectedNintexUser ? (
          <div style={{ padding: '10px 0', display: 'flex', flexDirection: 'column', gap: '12px' }}>
            <Persona
              text={`${selectedNintexUser.firstName || ''} ${selectedNintexUser.lastName || ''}`.trim() || selectedNintexUser.email || selectedUserId}
              secondaryText={selectedNintexUser.email}
              tertiaryText={`ID: ${selectedNintexUser.id}`}
              size={PersonaSize.size56}
            />
            <div style={{ display: 'grid', gridTemplateColumns: '110px 1fr', gap: '8px 12px', fontSize: '13px', marginTop: '10px', backgroundColor: '#f3f2f1', padding: '12px 16px', borderRadius: '4px', alignItems: 'center' }}>
              <strong>First Name:</strong> <OverflowText text={selectedNintexUser.firstName || '-'} />
              <strong>Last Name:</strong> <OverflowText text={selectedNintexUser.lastName || '-'} />
              <strong>Email:</strong> <OverflowText text={selectedNintexUser.email || '-'} />
              <strong>User ID:</strong> <OverflowText text={selectedNintexUser.id || '-'} style={{ fontFamily: 'monospace' }} />
            </div>
          </div>
        ) : null}
        <DialogFooter>
          <DefaultButton onClick={() => setIsUserDetailOpen(false)} text="Close" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};
